# ConfAutomation
## Automation for Zoom + OBS in Educational Environments
**Michael Lyle, Oakwood School, Morgan Hill, California**
------------------------------------------------------------------

This is software for Windows to launch OBS and Zoom together, and manage the gallery view and mouse of a computer.  It allows OBS and Zoom to be used together to deliver synchronous educational content.

It is a key component described in the paper "Integration of a Novel Multimedia Classroom for Distance and Hybrid Learning," where we share our approach of using OBS to mix down multiple classroom microphones and select from multiple classroom sources of video for hybrid learning (where some students are present in the classroom and some are remote).

The learning curve for OBS is steep.  With careful configuration of OBS and this software, we've endeavored to simplify usage for educators and integrate a low-cost solution that allows educators to seamlessly deliver existing lesson plans and pedagogical practices in distance and hybrid learning scenarios.

This software is written in Python.  PyInstaller and InnoSetup are used to produce an exectuable installer that can run on systems without Python installed.

The novel portions of this software are licensed under a permissive (0BSD) license, although pyhk3 is included which is licensed under the GPL version 2 (or any later version).

Below are the tasks that this automation performs:

- Minimizes its own console window (but keeps it around in case troubleshooting is required).
- Makes sure that the executables we need exists.
- Checks whether all monitors are present.
- Terminates any existing Zoom/OBS process
- Copies a template OBS profile into place to prevent inadvertent changes
- Starts OBS
- Starts Zoom
- Waits for conference to start and positions Zoom windows
  - Gallery on auxiliary screen
  - Chat / participants pop-out on laptop screen.
- Waits in a loop for hotkey button presses
  - CTRL-ALT-M: reposition mouse to known position (eases management of multimonitor views.
  - CTRL-ALT-G: moves the gallery to the next monitor.
  - CTRL-SHIFT-Q: quits
- Forcibly quits OBS & Zoom 
