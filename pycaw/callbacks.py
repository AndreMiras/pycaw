from ctypes import pointer

from comtypes import COMObject

from pycaw.api.audiopolicy import (
    IAudioSessionControl2,
    IAudioSessionEvents,
    IAudioSessionNotification,
)
from pycaw.api.endpointvolume import IAudioEndpointVolumeCallback
from pycaw.utils import AudioSession


class AudioSessionNotification(COMObject):
    """
    Helper for audio session created callbacks.

    Note
    ----
    In order for the AudioSessionNotification to work you need to play nicely
    by following these Windows rules:
    1.  Com needs to be in MTA. That is archived by defining
        the following flag before pycaw or comtypes are imported:
            sys.coinit_flags = 0
    2.  Get the AudioSessionManager:
            mgr = AudioUtilities.GetAudioSessionManager()
    3.  Create and register callback:
            MyCustomCallback(AudioSessionNotification):
                def on_session_created(self, new_session):
                    print("on_session_created")
            callback = MyCustomCallback()
            mgr.RegisterSessionNotification(callback)
    4.  Call the session enumerator (otherwise on_session_created wont work)
            mgr.GetSessionEnumerator()
    5.  Unregister, when you are finished:
            mgr.UnregisterSessionNotification(callback)

    Methods
    -------
    Override the following method:

    def on_session_created(self, new_volume, new_mute, event_context):
        Is fired, when a new audio session is created.
            new_session : pycaw.utils.AudioSession
    """

    _com_interfaces_ = (IAudioSessionNotification,)

    def OnSessionCreated(self, new_session):
        ctl2 = new_session.QueryInterface(IAudioSessionControl2)
        new_session = AudioSession(ctl2)
        self.on_session_created(new_session)

    def on_session_created(self, new_session):
        """pycaw user interface"""
        raise NotImplementedError


class AudioSessionEvents(COMObject):
    """
    Helper for audio session callbacks.

    Methods
    -------
    Override the following method(s):

    def on_simple_volume_changed(self, new_volume, new_mute, event_context):
        Is fired, when the audio session volume/mute changed.
            new_volume : float
                in range(0, 1)
            new_mute : int
                0, 1
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents

    def on_state_changed(self, new_state, new_state_id):
        Is fired, when the audio session state changed.
            new_state : str
                "Inactive", "Active", "Expired"
            new_state_id : int
                0, 1, 2

    def on_session_disconnected(self, disconnect_reason, disconnect_reason_id):
        Is fired, when the audio session disconnected "hard".
            Mostly on_state_changed == "Expired" is what you are looking for.
            see self.AudioSessionDisconnectReason for disconnect_reason
            The use is similar to on_state_changed
    """

    _com_interfaces_ = (IAudioSessionEvents,)

    # ======= DECODE RETURNED INT VALUE =======
    # see audiosessiontypes.h and audiopolicy.h
    AudioSessionState = ("Inactive", "Active", "Expired")

    AudioSessionDisconnectReason = (
        "DeviceRemoval",
        "ServerShutdown",
        "FormatChanged",
        "SessionLogoff",
        "SessionDisconnected",
        "ExclusiveModeOverride",
    )

    def OnSimpleVolumeChanged(self, new_volume, new_mute, event_context):
        self.on_simple_volume_changed(new_volume, new_mute, event_context)

    def OnStateChanged(self, new_state_id):
        new_state = self.AudioSessionState[new_state_id]
        self.on_state_changed(new_state, new_state_id)

    def OnSessionDisconnected(self, disconnect_reason_id):
        disconnect_reason = self.AudioSessionDisconnectReason[disconnect_reason_id]
        self.on_session_disconnected(disconnect_reason, disconnect_reason_id)

    def on_simple_volume_changed(self, new_volume, new_mute, event_context):
        """pycaw user interface"""
        pass

    def on_state_changed(self, new_state, new_state_id):
        """pycaw user interface"""
        pass

    def on_session_disconnected(self, disconnect_reason, disconnect_reason_id):
        """pycaw user interface"""
        pass


class AudioEndpointVolumeCallback(COMObject):
    """
    Helper for audio device callbacks.

    Methods
    -------
    Override the following method:

    def on_notify(self, new_volume, new_mute, event_context,
                  channels, channel_volumes):
        Is fired, when the audio device volume/mute changed.
            new_volume : float
                in range(0, 1)
            new_mute : int
                0, 1
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents
            channels : int
                count of channels
            channel_volumes : list : float
                the channel volumes in range(0, 1)
                len(channel_volumes) == channels
    """

    _com_interfaces_ = (IAudioEndpointVolumeCallback,)

    def OnNotify(self, pNotify):
        """Fired by Windows, when the audio device volume/mute changed"""

        # get the data of the PAUDIO_VOLUME_NOTIFICATION_DATA Structure
        notify_data = pNotify.contents

        channels = notify_data.nChannels
        # _.afChannelVolumes is a c_float_Array_8 -> convert to list
        channel_volumes = list(notify_data.afChannelVolumes)
        # remove from 8 value list everything out of channel range
        channel_volumes = channel_volumes[:channels]

        event_context = pointer(notify_data.guidEventContext)

        self.on_notify(
            notify_data.fMasterVolume,
            notify_data.bMuted,
            event_context,
            channels,
            channel_volumes,
        )

    def on_notify(self, new_volume, new_mute, event_context, channels, channel_volumes):
        """pycaw user interface"""
        raise NotImplementedError
