# Conference Automation System for OBS and Zoom
# Copyright (C) 2020 Michael P. Lyle

# Permission to use, copy, modify, and/or distribute this software for any 
# purpose with or without fee is hereby granted.

# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.

from pywinauto import Desktop, keyboard

import win32api, win32event, win32con
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

smallidx = 0

def copy_over(source, dest):
    """Copy a path from source to dest, making a backup.

    Keyword arguments:
    source -- the source path
    dest -- the dest path, including the destination directory name
    """
    import shutil

    bak_path = dest + '-bak'

    try:
        shutil.rmtree(bak_path)
        print("removed %s"%(bak_path))
    except Exception:
        pass

    try:
        print("moving %s to %s"%(dest, bak_path))
        shutil.move(dest, bak_path)
        print("moved %s to %s"%(dest, bak_path))
    except Exception:
        pass

    print("copying %s to %s"%(source, dest))
    shutil.copytree(source, dest)
    print("Completed copy.")

def copy_obs_profile():
    """Copy the OBS profile from our installation directory to %APPDATA%"""
    import sys

    dest_path = pathlib.Path(winshell.application_data(), "obs-studio")

    master_path = pathlib.Path(winshell.application_data(), "obs-master")

    if master_path.exists():
        print("Copying profile from master path")
        copy_over(str(master_path), str(dest_path))
        return

    prog_path = pathlib.Path(sys.argv[0]).parent.joinpath("obs-studio")
    if not prog_path.exists():
        print("Skipped copying obs-studio; no master copy and no copy in distribution")
        return

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

def get_zoom_pid():
    procs = find_procs_by_name('Zoom')
    if len(procs) < 1:
        raise Exception("No Zoom Process")
    return procs[0].pid

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
        try:
            txt = w.window_text().lower()
            if "confautomation" in txt:
                # Don't minimize the development environment!
                if "visual studio" in txt:
                    continue
                w.minimize()
        except Exception:
            pass

def check_really_exist_and_visible(specifications):
    try:
        for spec in specifications:
            if not spec.exists():
                return False
            if not spec.is_visible():
                return False
    except Exception:
        return False

    return True

def pop_out_zoom_controls(send_fullscreen=False):
    """Find the Zoom Meeting window, and type keys that pop out key windows"""
    # Use a desktop handle that's not UIA to move windows, because UIA doesn't work for some reason
    desktop = Desktop()

    # Be specific, match window name exactly.  Because fuzzy matching gets
    # the wrong window ("Zoom" or "Zoom Cloud Meetings")
    zoom = Desktop(backend="uia").window(title_re = '^Zoom Meeting$')
    chat = desktop.window(title_re = '^(Zoom Group )?Chat$')

    retries = 5

    while not zoom.exists(timeout=0):
        print("Waiting for zoom window")
        time.sleep(0.25)

    # Don't mess with window right after it appears, in hopes of fixing "Join Meeting"
    time.sleep(2)
    
    if send_fullscreen:
        zoom.type_keys('%f')

    print("Looking for participants and chat")
    while not check_really_exist_and_visible([desktop.participants, chat]):
        print("They don't exist, trying...")
        retries -= 1

        if retries <= 0:
            print("Reached retry limit")
            raise Exception("Could not pop out participants and chat")

        print("Waiting for popped-out participants & chat")
        if zoom.ContentRightPanel.exists(timeout=0) and zoom.ContentRightPanel.Participants.exists(timeout=0):
            print("Doing the popping ourselves of participants")
            zoom.set_focus()
            participants = zoom.ContentRightPanel.Participants
            participants.click_input(coords=(12,10))
            zoom.Pop_Out.click_input()
        else:
            zoom.type_keys('%u')

        if zoom.ContentRightPanel.exists(timeout=0) and zoom.ContentRightPanel.Chat_Expanded.exists(timeout=0):
            print("Doing the popping ourselves of chat")
            zoom.set_focus()
            chat_panel = zoom.ContentRightPanel.Chat_Expanded
            chat_panel.click_input(coords=(27,25))
            zoom.Pop_Out.click_input()
            print("pop-cycle complete")
        else:
            zoom.type_keys('%h')
        
        time.sleep(0.2)

    print("Moving participants window")
    desktop.participants.move_window(30,10)

    print("Moving chat window")
    chat.move_window(30,460)

def get_smallest_monitor():
    """Scans the monitor array, and returns the lowest resolution monitor"""
    smallest = 999999

    global smallidx

    smallidx = 0

    for i in range(len(monitors)):
        mon_size = monitors[i][2][2] - monitors[i][2][0]

        if mon_size < smallest:
            smallest = mon_size
            smallidx = i
            print("Selected %d as smallest mon, size %d"%(smallidx,smallest))

    return smallidx

def wait_for_key_up(keys):
    for key in keys:
        while win32api.GetAsyncKeyState(key) < 0:
            time.sleep(0.1)

def move_gallery_to_monitor(num):
    """Moves the gallery to the monitor index specified by num"""
    print("Getting desktop object")
    desktop = Desktop()
    print("Enumerating windows")
    windows = desktop.windows()
    print("searching...")
    mon_dims = monitors[num][2]

    for w in windows:
        try:
            if (w.window_text() == "Zoom Meeting"):
                print("Beginning gallery move")
                print(w.client_rect())

                # Leave room for taskbar, because we can't ensure that the zoom window is on top of it
                if num == 0:
                    taskbar_offset = 36
                else:
                    # Don't leave space on secondary monitors
                    taskbar_offset = 1
                target=(mon_dims[0]+1, mon_dims[1]+1, mon_dims[2]-mon_dims[0]-2, mon_dims[3]-mon_dims[1]-taskbar_offset)
                print(target)
                w.move_window(*target)
                w.set_focus()
                time.sleep(0.2)
                w.move_window(*target)

                print(w.client_rect())
        except Exception:
            print("Exception on window name, continuing")

    print("Completed gallery move, setting arm_time")

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

    try:
        minimize_ourselves()
    except Exception:
        pass

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

def key_move_meeting():
    """Macro key handler: move gallery view to next monitor"""
    global mon

    # Waiting for key release is necessary for well defined behavior, because
    # when we send macro presses, a key being still down can be harmful.
    print("key_move_meeting: Waiting for key release")
    wait_for_key_up([ord('G'), win32con.VK_LCONTROL, win32con.VK_RCONTROL, win32con.VK_LMENU, win32con.VK_RMENU])
    print("Keys released")

    mon = mon + 1

    if mon >= len(monitors):
        mon = 0

    print("moving gallery to monitor %d"%(mon))
    move_gallery_to_monitor(mon)

def key_move_meeting_C():
    """Macro key handler: move gallery view to tertiary monitor"""
    global mon

    # Waiting for key release is necessary for well defined behavior, because
    # when we send macro presses, a key being still down can be harmful.
    print("key_move_meeting_C: Waiting for key release")
    wait_for_key_up([ord('C'), win32con.VK_LCONTROL, win32con.VK_RCONTROL, win32con.VK_LMENU, win32con.VK_RMENU])
    print("Keys released")

    mon = smallidx

    if mon >= len(monitors):
        return

    print("moving gallery to monitor %d"%(mon))
    move_gallery_to_monitor(mon)

def key_move_meeting_L():
    """Macro key handler: move gallery view to laptop screen"""
    global mon

    # Waiting for key release is necessary for well defined behavior, because
    # when we send macro presses, a key being still down can be harmful.
    print("key_move_meeting_L: Waiting for key release")
    wait_for_key_up([ord('L'), win32con.VK_LCONTROL, win32con.VK_RCONTROL, win32con.VK_LMENU, win32con.VK_RMENU])
    print("Keys released")

    mon = 0

    if mon >= len(monitors):
        return

    print("moving gallery to monitor %d"%(mon))
    move_gallery_to_monitor(mon)

def key_pop_out_zoom():
    wait_for_key_up([ord('Z'), win32con.VK_LCONTROL, win32con.VK_RCONTROL, win32con.VK_LMENU, win32con.VK_RMENU])
    try:
        pop_out_zoom_controls()
    except Exception:
        pass

def key_mute_zoom():
    """Send the mute keypress to Zoom"""
    print("key_move_meeting: Waiting for key release")
    wait_for_key_up([win32con.VK_SPACE, win32con.VK_LCONTROL, win32con.VK_RCONTROL, win32con.VK_LMENU, win32con.VK_RMENU])
    print("Keys released")

    desktop = Desktop()

    zoom = desktop.window(title_re = '^Zoom Meeting$')

    # Toggles mute
    zoom.type_keys('%a')

def key_center_mouse():
    """Moves the mouse to 500,500; somewhere in the middle of main display"""
    win32api.SetCursorPos((500,500))

def check_already_running():
    global mutex

    mutex = win32event.CreateMutex(None, False, "Oakwood_confautomation")
    if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
        show_warning("There is already another copy of ConfAutomation running; please exit it.")
        import sys
        os._exit(1)

def pyhk_go():
    hot = pyhk3.pyhk()

    # add hotkeys.  Additionally, CTRL-SHIFT-Q exits (built into pyhk3)
    id1 = hot.addHotkey(['Ctrl', 'Alt', 'G'], key_move_meeting, isThread=True)
    id5 = hot.addHotkey(['Ctrl', 'Alt', 'L'], key_move_meeting_L, isThread=True)
    id6 = hot.addHotkey(['Ctrl', 'Alt', 'C'], key_move_meeting_C, isThread=True)

    id2 = hot.addHotkey(['Ctrl', 'Alt', 'M'], key_center_mouse)
    id3 = hot.addHotkey(['Ctrl', 'Alt', 'Space'], key_mute_zoom, isThread=True)
    id4 = hot.addHotkey(['Ctrl', 'Alt', 'Z'], key_pop_out_zoom, isThread=True)

    print("Waiting for hotkeys in pyhk_go")
    hot.start()

    kill_procs_by_name('Zoom')
    kill_procs_by_name('OBS')
    import sys
    os._exit(0)

def main():
    """Main function: starts conference & waits for hotkeys, coordinates shutdown"""
    #check_already_running()

    global monitors

    monitors = win32api.EnumDisplayMonitors()
    print(monitors)

    conference_start()

    global mon

    mon = get_smallest_monitor()

    # Wait for user to start meeting, pop out controls
    try:
        pop_out_zoom_controls(send_fullscreen=True)
    except Exception:
        pass

    move_gallery_to_monitor(mon)
    pyhk_go()

if __name__ == "__main__":
    # execute only if run as a script
    main()