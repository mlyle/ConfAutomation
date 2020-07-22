# C:\Users\ddennis\AppData\Local\Programs\Python\Python38\Scripts\pyinstaller.exe --noconfirm testing.spec

from pywinauto import Desktop, keyboard

import win32api
import time

monitors = win32api.EnumDisplayMonitors()
print(monitors)

def pop_out_zoom_controls():
    desktop = Desktop()
    zoom = desktop.Zoom_Meeting

    zoom.type_keys('%h')
    zoom.type_keys('%u')
    desktop.participants.move_window(30,30)
    desktop.chat.move_window(200,30)

# Retry getting the zoom meeting window, because it's finicky
for i in range(3):
    try:
        pop_out_zoom_controls()
        break
    except Exception:
        time.sleep(0.5)

smallest=999999
mon = 0

for i in range(len(monitors)):
    mon_dims = monitors[i][2]
    if mon_dims[2] < smallest:
        smallest = mon_dims[2]
        mon = i

def move_gallery_to_monitor(num):    
    desktop = Desktop()
    windows = desktop.windows()

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

move_gallery_to_monitor(mon)

import pyhk3

def key_move_meeting():
    global mon

    mon = mon + 1

    if mon >= len(monitors):
        mon = 0

    move_gallery_to_monitor(mon)

hot = pyhk3.pyhk()
print(hot.getHotkeyListNoSingleNoModifiers())

# add hotkey
id1 = hot.addHotkey(['Ctrl', 'Alt', 'G'], key_move_meeting)

hot.start()