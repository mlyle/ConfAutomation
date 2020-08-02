from pywinauto import Desktop, keyboard

import win32api, win32event
from winerror import ERROR_ALREADY_EXISTS
import time

import psutil
import winshell
import os
import pathlib
import pyhk3

# XXX these hardcoded paths are unfortunate
path_zoom = pathlib.Path(winshell.application_data(), 'Zoom', 'bin', 'zoom.exe')
path_obs = pathlib.Path('C:\\Program Files\\obs-studio\\bin\\64bit\\obs64.exe')

gallery_arm_time = 0

def copy_over(source, dest):
    """Copy a path from source to dest, making a backup.

    Keyword arguments:
    source -- the source path
    dest -- the dest path, including the destination directory name
    """
    import shutil

    try:
        shutil.rmtree(dest + "-bak")
        print("Moved old obs-studio directory to backup")
    except Exception:
        pass
    shutil.move(dest, dest + "-bak")
    shutil.copytree(source, dest)
    print("Completed copying obs-studio configuration directory")

def copy_obs_profile():
    """Copy the OBS profile from our installation directory to %APPDATA%"""
    import sys

    prog_path = pathlib.Path(sys.argv[0]).parent.joinpath("obs-studio")
    if not prog_path.exists():
        print("Skipped copying obs-studio because it's missing in distribution")
        return

    dest_path = str(pathlib.Path(winshell.application_data(), "obs-studio"))
    copy_over(str(prog_path), str(dest_path))

def show_warning(text):
    """Open a warning dialog box, allowing the user to choose to exit"""
    if win32api.MessageBox(0, "Warning: " + text, 'ConfAutomation', 0x1031) != 1:
        import sys
        sys.exit()

def ensure_exists(path):
    """Warn if a path doesn't exist"""
    if not path.exists():
        show_warning("Could not find %s which is REQUIRED FOR OPERATION"%(str(path)))

def find_procs_by_name(name):
    """Return a list of processes matching 'name'."""
    ls = []
    for p in psutil.process_iter(['name']):
        psname = p.info['name']

        if psname is None:
            continue

        if name.lower() in psname.lower():
            ls.append(p)
    return ls

def kill_procs_by_name(name, noisy=False):
    """Forcibly/immediately terminate processes matching name"""
    first = True
    procs = find_procs_by_name(name)
    for proc in procs:
        if noisy and first:
            show_warning("%s is already running; will stop and restart it!" % name)
        first = False
        proc.kill()

def start_zoom():
    """Launch Zoom process"""
    os.startfile(str(path_zoom))

def start_obs():
    """Start OpenBroadcastSystem"""
    oldpath = os.getcwd()

    os.chdir(str(path_obs.parent))
    os.startfile(path_obs)
    os.chdir(oldpath)

def minimize_ourselves():
    """Find our console window and minimize it"""
    desktop = Desktop()
    windows = desktop.windows()

    for w in windows:
        txt = w.window_text().lower()
        if "confautomation" in txt:
            # Don't minimize the development environment!
            if "visual studio" in txt:
                continue
            w.minimize()

def pop_out_zoom_controls():
    """Find the Zoom Meeting window, and type keys that pop out key windows"""
    desktop = Desktop()
    zoom = desktop.Zoom_Meeting

    # This sends the keys to open the participants and chats lists.
    # They must already be selected as "popped out" in Zoom
    zoom.type_keys('%h')
    zoom.type_keys('%u')
    desktop.participants.move_window(30,30)
    desktop.chat.move_window(200,30)

def select_smallest_monitor():
    """Scans the monitor array, and finds the lowest resolution monitor"""
    # XXX should return smallest mon idx instead of setting mon
    smallest=999999
    global mon

    mon = 0

    for i in range(len(monitors)):
        mon_dims = monitors[i][2]
        if mon_dims[2] < smallest:
            smallest = mon_dims[2]
            mon = i

def move_gallery_to_monitor(num):
    """Moves the gallery to the monitor index specified by num"""
    global gallery_arm_time

    now = time.time()

    interval = now - gallery_arm_time

    # Moving windows onto the high resolution display takes some time.
    # Only be willing to answer a keystroke if it's been half a second since
    # when we finished processing the previous keystroke.
    if interval < 0.5:
        return

    desktop = Desktop()
    windows = desktop.windows()
    mon_dims = monitors[num][2]

    for w in windows:
        if (w.window_text() == "Zoom Meeting"):
            print(w.client_rect())
            target=(mon_dims[0]+1, mon_dims[1]+1, mon_dims[2]-mon_dims[0]-2, mon_dims[3]-mon_dims[1]-2)
            print(target)
            w.move_window(*target)
            w.set_focus()
            w.move_window(*target)

            print(w.client_rect())

            time.sleep(0.5)

            desktop.Zoom_Meeting.type_keys('%f')

    gallery_arm_time = time.time()

def conference_start():
    """Performs key conference start activities.

    - Makes sure that the executables we need exists.
    - Gathers a monitor array and checks for sanity
    - Terminates any existing Zoom/OBS process
    - Copies the OBS profile into place
    - Starts OBS
    - Starts Zoom
    - Waits for conference to start and positions Zoom windows
    """
    ensure_exists(path_zoom)
    ensure_exists(path_obs)

    global monitors

    monitors = win32api.EnumDisplayMonitors()
    print(monitors)

    if len(monitors) != 3:
        show_warning("Expected 3 monitors but found %d"%(len(monitors)))

    # Ensure that Zoom & OBS are not running
    kill_procs_by_name('Zoom', True)
    kill_procs_by_name('OBS', True)

    # copy in OBS profile directory from software distribution--
    # iff it exists.
    copy_obs_profile()

    start_obs()

    # Unfortunately, with OBS starting minimized there is not a wonderful way
    # to know when it has completed launching.  Instead, we wait a few seconds.
    # There's no grave consequence if Zoom comes up first, though it is nice
    # to make this as deterministic as possible.
    time.sleep(4.5)
    start_zoom()
    try:
        minimize_ourselves()
    except Exception:
        pass

    # Wait for user to start meeting, pop out controls
    # XXX the program is in an unfortunate "limbo" during this interval;
    # I worry about user launching multiple times, etc.
    while True:
        try:
            pop_out_zoom_controls()
            break
        except Exception:
            print("Failed to move zoom controls; trying again...")
            time.sleep(0.25)

    select_smallest_monitor()

    move_gallery_to_monitor(mon)

def key_move_meeting():
    """Macro key handler: move gallery view to next monitor"""
    global mon

    mon = mon + 1

    if mon >= len(monitors):
        mon = 0

    print("moving gallery to monitor %d"%(mon))
    move_gallery_to_monitor(mon)

def key_center_mouse():
    """Moves the mouse to 500,500; somewhere in the middle of main display"""
    win32api.SetCursorPos((500,500))

def check_already_running():
    global mutex

    mutex = win32event.CreateMutex(None, False, "Oakwood_confautomation")
    if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
        show_warning("There is already another copy of ConfAutomation running; please exit it.")
        import sys
        sys.exit(1)

def main():
    """Main function: starts conference & waits for hotkeys, coordinates shutdown"""
    check_already_running()

    conference_start()
    hot = pyhk3.pyhk()

    # add hotkeys.  Additionally, CTRL-SHIFT-Q exits (built into pyhk3)
    id1 = hot.addHotkey(['Ctrl', 'Alt', 'G'], key_move_meeting)
    id2 = hot.addHotkey(['Ctrl', 'Alt', 'M'], key_center_mouse)

    hot.start()

    kill_procs_by_name('Zoom')
    kill_procs_by_name('OBS')

if __name__ == "__main__":
    # execute only if run as a script
    main()