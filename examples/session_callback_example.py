"""
Enter the right 'app_name' and play something to initiate a WASAPI session.
Then launch this file in interactive mode:

python -i session_callback_example.py

following "IAudioSessionEvents" callbacks are supported:
IAudioSessionEvents.OnSimpleVolumeChanged()         Gets called on volume and mute change
IAudioSessionEvents.OnStateChanged()                Gets called on session state change (active/inactive/expired)
IAudioSessionEvents.OnSessionDisconnected()         Gets called on for example Speaker unplug

https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessionevents

"""

from __future__ import print_function
from comtypes import COMObject, COMError
from pycaw.pycaw import (AudioUtilities, IAudioSessionEvents)


app_name="msedge.exe"


class AudioSessionEvents(COMObject):
    _com_interfaces_ = [IAudioSessionEvents]

    def __init__(self):
        # not necessary only to decode the returned int value
        # see mmdeviceapi.h and audiopolicy.h
        self.AudioSessionState = ["AudioSessionStateInactive", "AudioSessionStateActive", "AudioSessionStateExpired"]
        self.AudioSessionDisconnectReason = ["DisconnectReasonDeviceRemoval", "DisconnectReasonServerShutdown", "DisconnectReasonFormatChanged", "DisconnectReasonSessionLogoff", "DisconnectReasonSessionDisconnected", "DisconnectReasonExclusiveModeOverride"]

    def OnSimpleVolumeChanged(self, NewVolume, NewMute, EventContext):
        print(':: OnSimpleVolumeChanged callback')
        print(f"NewVolume: {NewVolume}; NewMute: {NewMute}; EventContext: {EventContext}")

    def OnStateChanged(self, NewState):
        print(':: OnStateChanged callback')
        translate = self.AudioSessionState[NewState]
        print(translate)

    def OnSessionDisconnected(self, DisconnectReason):
        print(':: OnSessionDisconnected callback')
        translate = self.AudioSessionDisconnectReason[DisconnectReason]
        print(translate)


try:
    sessions = AudioUtilities.GetAllSessions()
except COMError:
    exit("No speaker set up")

# grap the right session (this is only for trying out, but with multiple matching sessions it will only grap the last)

your_app_session = None
for session in sessions:
    if session.Process and session.Process.name() == app_name: # app_name = "msedge.exe"
        your_app_session = session

if not your_app_session:
    exit("Enter the right 'app_name', start it and play something")


# Adding the callback by accessing the IAudioSessionControl2 interface through ._ctl

callback = AudioSessionEvents()
your_app_session._ctl.RegisterAudioSessionNotification(callback)

print("Ready to go!")
print("Change the volume / mute state / close the app or unplug your speaker")
print("and watch the callbacks ;)\n")