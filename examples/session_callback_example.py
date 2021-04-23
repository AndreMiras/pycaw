"""
Enter the right 'app_name' and play something to initiate a WASAPI session.
Then launch this file ;)

Requirements: Python >= 3.6 - f strings ;)

following "IAudioSessionEvents" callbacks are supported:
:: Gets called on volume and mute change:
IAudioSessionEvents.OnSimpleVolumeChanged()
:: Gets called on session state change (active/inactive/expired)
IAudioSessionEvents.OnStateChanged()
:: Gets called on for example Speaker unplug
IAudioSessionEvents.OnSessionDisconnected()

https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessionevents

"""
import time

from comtypes import COMError, COMObject

from pycaw.pycaw import AudioUtilities, IAudioSessionEvents

app_name = "msedge.exe"


class AudioSessionEvents(COMObject):
    _com_interfaces_ = [IAudioSessionEvents]

    def __init__(self):
        # not necessary only to decode the returned int value
        # see mmdeviceapi.h and audiopolicy.h
        self.AudioSessionState = [
            "AudioSessionStateInactive",
            "AudioSessionStateActive",
            "AudioSessionStateExpired"
        ]
        self.AudioSessionDisconnectReason = [
            "DisconnectReasonDeviceRemoval",
            "DisconnectReasonServerShutdown",
            "DisconnectReasonFormatChanged",
            "DisconnectReasonSessionLogoff",
            "DisconnectReasonSessionDisconnected",
            "DisconnectReasonExclusiveModeOverride"
        ]

    def OnSimpleVolumeChanged(self, NewVolume, NewMute, EventContext):
        print(':: OnSimpleVolumeChanged callback')
        print(f"NewVolume: {NewVolume};"
              "NewMute: {NewMute};"
              "EventContext: {EventContext.contents}")

    def OnStateChanged(self, NewState):
        print(':: OnStateChanged callback')
        translate = self.AudioSessionState[NewState]
        print(translate)

    def OnSessionDisconnected(self, DisconnectReason):
        print(':: OnSessionDisconnected callback')
        translate = self.AudioSessionDisconnectReason[DisconnectReason]
        print(translate)


def add_callback(app_name):
    try:
        sessions = AudioUtilities.GetAllSessions()
    except COMError:
        exit("No speaker set up")

    # grap the right session
    app_found = False
    for session in sessions:
        if session.Process and session.Process.name() == app_name:

            app_found = True
            callback = AudioSessionEvents()
            # Adding the callback by accessing the
            # IAudioSessionControl2 interface through ._ctl
            session._ctl.RegisterAudioSessionNotification(callback)

    if not app_found:
        exit("Enter the right 'app_name', start it and play something")

    print("Ready to go!")
    print("Change the volume / mute state"
          "/ close the app or unplug your speaker")
    print("and watch the callbacks ;)\n")

    try:
        # wait 300 seconds for callbacks
        time.sleep(300)
    except KeyboardInterrupt:
        pass
    finally:
        print("\nTschüss")


if __name__ == "__main__":
    add_callback(app_name)
