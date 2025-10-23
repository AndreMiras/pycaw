"""
Enter the right 'app_name' and play something to initiate a WASAPI session.
Then launch this file ;)

Requirements: Python >= 3.6 - f strings ;)

following "IAudioSessionEvents" callbacks are supported:
:: Gets called on volume and mute change:
IAudioSessionEvents.OnSimpleVolumeChanged()
-> on_simple_volume_changed()

:: Gets called on session state change (active/inactive/expired):
IAudioSessionEvents.OnStateChanged()
-> on_state_changed()

:: Gets called on for example Speaker unplug
IAudioSessionEvents.OnSessionDisconnected()
-> on_session_disconnected()

https://docs.microsoft.com/en-us/windows/win32/api/audiopolicy/nn-audiopolicy-iaudiosessionevents

"""

import time

from comtypes import COMError

from pycaw.callbacks import AudioSessionEvents
from pycaw.utils import AudioUtilities

app_name = "msedge.exe"


class MyCustomCallback(AudioSessionEvents):
    def on_simple_volume_changed(self, new_volume, new_mute, event_context):
        print(
            ":: OnSimpleVolumeChanged callback\n"
            f"new_volume: {new_volume}; "
            f"new_mute: {new_mute}; "
            f"event_context: {event_context.contents}"
        )

    def on_state_changed(self, new_state, new_state_id):
        print(
            ":: OnStateChanged callback\n" f"new_state: {new_state}; id: {new_state_id}"
        )

    def on_session_disconnected(self, disconnect_reason, disconnect_reason_id):
        print(
            ":: OnSessionDisconnected callback\n"
            f"disconnect_reason: {disconnect_reason}; "
            f"id: {disconnect_reason_id}"
        )


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
            callback = MyCustomCallback()
            # Adding the callback
            session.register_notification(callback)

    if not app_found:
        exit("Enter the right 'app_name', start it and play something")

    print("Ready to go!")
    print("Change the volume / mute state / close the app or unplug your speaker")
    print("and watch the callbacks ;)\n")

    try:
        # wait 300 seconds for callbacks
        time.sleep(300)
    except KeyboardInterrupt:
        pass
    finally:
        # unregister callback(s)
        # unregister_notification()
        # (only if it was also registered.)
        # pycaw.utils -> unregister_notification()
        for session in sessions:
            session.unregister_notification()

        print("\nTschüss")


if __name__ == "__main__":
    add_callback(app_name)
