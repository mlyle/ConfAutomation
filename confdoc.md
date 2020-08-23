Conference System Deployment Documentation
==========================================

Laptop Setup Procedure
----------------------

-   If not already installed, install OBS and OBSVirtualCam.

-   As administrator, run VBCableSetupx64.exe from the VB Audio Cable
    > setup. (This provides the audio path from OBS -\> Zoom)

-   As administrator, run FL2000Setup. (This makes the USB to HDMI
    > adapter on the cart work)

-   Install Zoom

-   Run CA-SETUP (This is the conference automation software that MPL
    > wrote)

    -   Check the box to create desktop shortcut

    -   Uncheck the box to launch after installation

-   Set Elmo switches on side to "HD Projector" and 720P.

-   Restart. Ensure TV/projector and displays on. Connect USB and HDMI
    > cable after reboot completes.

-   Arrange displays:

    -   Ensure displays are extended

    -   Set the monitor on the cart to 800x600 resolution

    -   Arrange displays to mirror how the cart monitor (3), laptop (1),
        > and TV/projector (2) are physically arranged in the classroom.

-   Configure the laptop to send sound to the classroom
    > television/projector.

-   Change laptop power management to sleep displays "after 5 hours"
    > when plugged in.

-   Launch ConfAutomation, which runs OBS & Zoom.

-   Setup UHF microphone (teacher mic) and receiver

    -   Ensure UHF microphone and receiver are paired.

    -   Make sure UHF microphone volume is all the way up (V20)

    -   Use both UHF microphone buttons to select the classroom's
        > channel

    -   Turn the microphone off, wait 10 seconds, and turn it back on to
        > allow the receiver to find the new channel.

-   Change OBS configuration

    -   File-\>Settings-\>Audio:

        -   Set Monitoring Destination as "Cable Input".

        -   Desktop Audio: set source of Desktop Audio to classroom
            > television

        -   Mic Audio properties: set source of Mic to USB Audio Device
            > (lapel mic)

        -   Mic2 Audio properties: set source of Mic2 to UM02
            > (conference microphone)

        -   OK

    -   In Audio Mixer panel: click gear next to audio source: Advanced
        > audio properties: Make sure Desktop Audio, Mic, and Mic2 are
        > all set to "Monitor and Output"

    -   In the Source panel, configure each of the video sources; click
        > on each scene, double click on audio source, confirm you can
        > get video from each of the appropriate sources.

        -   Classroom view set source to PTZ camera

        -   Elmo set source to Elmo (USB Video)

        -   Instructor laptop set source to integrated webcam

    -   Confirm that we can toggle between video sources, and that all 3
        > audio sources (Windows sound, lapel mic, area mic) make it
        > into OBS.

        -   CTRL-ALT-C

        -   CTRL-ALT-E

        -   CTRL-ALT-S

        -   Talk into teacher mic, watch Mic audio bar to see if its
            > picking up the sound

        -   Stand by the cart & talk, watch Mic2 audio bar to see if its
            > picking up the sound

    -   Right click on video view in OBS, select 'FullScreen Projector
        > Preview' and select classroom TV/projector. Verify that the
        > tv/projector is showing the video view that you see in OBS.

    -   Press CTRL-ALT-C to make sure the classroom view is showing on
        > the TV/projector.

    -   Minimize OBS window

-   Adjust Zoom configuration

    -   Choose "USB Audio Device" as "speakers", set level all the way
        > up.

    -   Choose "Cable Output" as "microphone", set level all the way up,
        > disable automatically adjust mic level

    -   Choose OBS-Cam in Zoom as the video source

    -   Check the "Enable HD" checkbox

-   Start a test Zoom conference.

    -   Ensure Zoom is set to send video and automatically join audio

    -   Pick gallery view on the test Zoom conference

    -   Pop out chat & participants

    -   Make sure view is on Classroom (CTRL-ALT-C) before exiting, so
        > that next time it starts on this view

    -   Do not end Zoom meeting

-   Important: Exit OBS with [File-Exit]{.ul} manually and [wait 5
    > seconds.]{.ul}

-   Hit CTRL-SHIFT-Q to exit automation and Zoom.

-   In Explorer, go to %APPDATA%, and rename the folder obs-studio to
    > obs-master

    -   This "saves" the OBS configuration for the automation to copy in
        > place on each startup.

    -   If you ever have to repeat this step, you'll have to erase the
        > old obs-master before renaming.

-   Start ConfAutomation again to test the setup end to end. Launch a
    > conference. Confirm all audio sources and video sources make it to
    > remote participants.

-   Program initial presets into camera; 1 "instructor view", 2 "wide
    > instructor view", 3 "classroom view"

**Note:** Sometimes it has been necessary to re-set the preview screen
in OBS for the change to "stick." If this is necessary, make the change,
then exit OBS, go to %AppData%, erase obs-master, and rename obs-studio
to obs-master (again).

Conference System Instructor Documentation
==========================================

Normal Startup Procedure
------------------------

1.  Restart your computer (this is good to do at least daily). USB and
    > HDMI cables should be unplugged when rebooting.

2.  Turn on classroom television.

3.  **After the computer has completed starting up,** plug in USB and
    > HDMI cables. Wait at least 15 seconds for the computer to settle.

    a.  Confirm that you can see video on all 3 monitors (laptop,
        > classroom TV, auxiliary monitor).

    b.  Confirm audio output is set to classroom TV in Sound Settings.

4.  Make sure the monitor and speakers on the cart are turned on.

5.  Unplug, turn on, & wear the teacher microphone.

6.  Double click on the "ConfAutomation" shortcut, and wait for Zoom to
    > start. (\~20 seconds)

7.  **If the view of the classroom shows on the incorrect monitor:** see
    > note below.

8.  If you have used Zoom outside your classroom, make sure in Zoom
    > settings that:

    a.  Camera is set to OBS-Camera in Video settings

    b.  In audio settings, Speaker is set to USB Audio Device and you
        > can hear the piano.

    c.  Microphone is set to "Cable Output", volume is all the way up,
        > and "Automatically Adjust Volume" is turned off.

9.  In Zoom, start or join a meeting.

10. If necessary, press CTRL-ALT-G to move the gallery view to the
    > monitor you want it on (usually this would be the monitor under
    > the camera). You can also use this to access Zoom controls.

11. When the meeting has completed, end the conference in Zoom. Then
    > hold down CTRL and SHIFT and press Q to shut down all conferencing
    > software.

12. Turn off the classroom television. Remember to charge your
    > microphone.

**If the view of the classroom shows up on the wrong monitor before the
conference starts:** This is slightly tricky: Move the mouse to where
the classroom view has appeared; right click; pick Fullscreen and choose
the "correct" monitor for the preview to show up on (projector or main
classroom TV). This is a bug that there wasn't time to fix before
deployment. Email
[[techhelp\@oakwoodway.org]{.ul}](mailto:techhelp@oakwoodway.org) to
report this problem.

**If video is blurry**: If the camera has not been moved in a long time
(\~2 hours), autofocus is disabled. Select a location preset to
re-enable autofocus.

**If ConfAutomation gets stuck**, it may be necessary to right-click on
its icon in the task bar and select "Close Window" before you are able
to launch it again. This mostly happens when you launch ConfAutomation
but never start a conference in Zoom.

**Chat must remain enabled** in your Zoom account (though it can be set
to only allow messages to the host) for now. When sharing your screen,
please make sure that sensitive content from chat (and other windows,
like Veracross) is covered by another window.

### Zoom Hotkeys

### Zoom's shortcuts only work by default when the Zoom window is selected (clicking on it on the side monitor is how you'd usually do this. The main useful shortcut is ALT-F2 to get back to the gallery view, but [[https://support.zoom.us/hc/en-us/articles/205683899-Hot-Keys-and-Keyboard-Shortcuts-for-Zoom]{.ul}](https://support.zoom.us/hc/en-us/articles/205683899-Hot-Keys-and-Keyboard-Shortcuts-for-Zoom) has a complete list of shortcuts.

### OpenBroadcastSystem Hotkeys

CTRL-ALT-C Send an entire **classroom** view\
CTRL-ALT-E Send the view from the **Elmo**\
CTRL-ALT-S Send the view from the main laptop **screen\
**(Make sure all sensitive content, including chat, is covered by what
you want to share)\
CTRL-ALT-P Send a **picture-in-picture** view with the Elmo and
Classroom view\
CTRL-ALT-L Send a view from the **laptop**'s built in camera\
CTRL-ALT-B **Blank** green screen (don't send video)

### Conference Automation Software Hotkeys

These are activated after the Zoom meeting has begun and the automation
has moved the gallery view to the auxiliary monitor.

CTRL-ALT-G Move the **gallery** of students to a different monitor\
This is useful if a remote student will be presenting to the classroom.\
CTRL-ALT-M Move the **mouse** to the lower middle of the laptop screen\
CTRL-SHIFT-Q Stops the conference and **quits** Zoom, OBS, and the
automation system.\
CTRL-ALT-SpaceBar Toggles whether we are muted in the Zoom meeting.
(Remember to unmute!)\
This is mostly useful when many classrooms join one meeting.

### Camera Presets

There are three preset views on the camera. The remote control is an
infrared control that you must point at the camera. When we configured
the classroom, we set them as:

1.  Instructor desk view

2.  Wide instructor desk view

3.  Overall classroom overview.

You can press each of these buttons to move the camera between views.
You can use the other controls on the remote to move the camera's view.
Holding a preset button for 5 seconds reprograms the location of that
preset. (I usually do this twice to make sure it "sticks").

Emergency Procedures
--------------------

If the automation isn't working, you can hold a conference without it.
The quality and flexibility will be degraded but it is a way to hold a
class until any problems are resolved.

1.  Make sure ConfAutomation isn't running. If necessary, right click on
    > its icon on the task bar and select "Close".

2.  Set the speaker in Zoom to USB Audio Device.

3.  Set the microphone to UM02. This is the conference microphone on the
    > cart.

4.  Set the camera in Zoom to Logitech PTZ Pro II.

5.  Start your meeting as you normally would, and use the Zoom controls.
