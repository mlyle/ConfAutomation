from pywinauto import Desktop, keyboard

desktop = Desktop()
windows = desktop.windows()

import win32api

monitors = win32api.EnumDisplayMonitors()
print(monitors)

def move_gallery_to_monitor(num):    
    for w in windows:
        mon_dims = monitors[num][2]
        if (w.window_text() == "Zoom Meeting"):
            print(w.client_rect())
            target=(mon_dims[0]+1, mon_dims[1]+1, mon_dims[2]-mon_dims[0]-2, mon_dims[3]-mon_dims[1]-2)
            print(target)
            w.move_window(*target)
            w.set_focus()
            w.move_window(*target)

            print(w.client_rect())

smallest=999999

mon = 0

for i in range(len(monitors)):
    mon_dims = monitors[i][2]
    if mon_dims[2] < smallest:
        smallest = mon_dims[2]
        mon = i

zoom = desktop.Zoom_Meeting

zoom.type_keys('%h')
zoom.type_keys('%u')

desktop.participants.move_window(30,30)
desktop.chat.move_window(200,30)
move_gallery_to_monitor(mon)

import pyhk3

def key_move_meeting():
    global mon
    desktop = Desktop()
    windows = desktop.windows()
    mon = mon + 1
    if mon >= len(monitors):
        mon = 0

    move_gallery_to_monitor(mon)

def key_quit():
    import sys
    sys.exit()

hot = pyhk3.pyhk()
print(hot.getHotkeyListNoSingleNoModifiers())
# add hotkey
id1 = hot.addHotkey(['Ctrl', 'Alt', 'G'], key_move_meeting)
id2 = hot.addHotkey(['Ctrl', 'Alt', '8'], key_quit)

hot.start()