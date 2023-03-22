from ctypes import pointer

from comtypes import COMObject

from pycaw.api.audiopolicy import (
    IAudioSessionControl2,
    IAudioSessionEvents,
    IAudioSessionNotification,
)
from pycaw.api.endpointvolume import IAudioEndpointVolumeCallback
from pycaw.api.mmdeviceapi import IMMNotificationClient
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

    def on_display_name_changed(self, new_display_name, event_context):
        Is fired, when the audio session name is changed.
            new_display_name : str
                The new name that is displayed.
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents.

    def OnIconPathChanged(self, new_icon_path, event_context):
        Is fired, when the audio session icon path is changed.
            new_icon_path : str
                The new path of the icon that is displayed.
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents.

    def on_simple_volume_changed(self, new_volume, new_mute, event_context):
        Is fired, when the audio session volume/mute changed.
            new_volume : float
                in range(0, 1)
            new_mute : int
                0, 1
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents.

    def OnChannelVolumeChanged(self, channel_count, new_channel_volume_array,
                                                changed_channel, event_context):
        Is fired, when the audio session channels volume changed.
            channel_count: int
                This parameter specifies the number of audio channels in the session
                submix.
            new_channel_volume_array : float array
                values in range(0, 1)
            changed_channel : int
                The number (x) of the channel whose volume level changed.
                Use (x-1) as index of new_channel_volume_array
                to get the new volume for the changed_channel (x)
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents.

    def OnGroupingParamChanged(self, new_grouping_param, event_context):
        Is fired, when the grouping parameter for the session has changed.
            new_grouping_param : comtypes.GUID
                points to a grouping-parameter GUID.
            event_context : comtypes.GUID
                the guid "should" be unique to who made the changes.
                access guid str with event_context.contents.

    def on_state_changed(self, new_state, new_state_id):
        Is fired, when the audio session state changed.
            new_state : str
                "Inactive", "Active", "Expired"
            new_state_id : int
                0, 1, 2

    def on_session_disconnected(self, disconnect_reason, disconnect_reason_id):
        Is fired, when the audio session disconnected "hard".
            Mostly on_state_changed == "Expired" is what you are looking for.
            see self.AudioSessionDisconnectReason for disconnect_reason.
            The use is similar to on_state_changed.
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

    def OnDisplayNameChanged(self, new_display_name, event_context):
        self.on_display_name_changed(new_display_name, event_context)

    def OnIconPathChanged(self, new_icon_path, event_context):
        self.on_icon_path_changed(new_icon_path, event_context)

    def OnSimpleVolumeChanged(self, new_volume, new_mute, event_context):
        self.on_simple_volume_changed(new_volume, new_mute, event_context)

    def OnChannelVolumeChanged(
        self, channel_count, new_channel_volume_array, changed_channel, event_context
    ):
        self.on_channel_volume_changed(
            channel_count, new_channel_volume_array, changed_channel, event_context
        )

    def OnGroupingParamChanged(self, new_grouping_param, event_context):
        self.on_grouping_param_changed(new_grouping_param, event_context)

    def OnStateChanged(self, new_state_id):
        new_state = self.AudioSessionState[new_state_id]
        self.on_state_changed(new_state, new_state_id)

    def OnSessionDisconnected(self, disconnect_reason_id):
        disconnect_reason = self.AudioSessionDisconnectReason[disconnect_reason_id]
        self.on_session_disconnected(disconnect_reason, disconnect_reason_id)

    def on_display_name_changed(self, new_display_name, event_context):
        """pycaw user interface"""
        pass

    def on_icon_path_changed(self, new_icon_path, event_context):
        """pycaw user interface"""
        pass

    def on_simple_volume_changed(self, new_volume, new_mute, event_context):
        """pycaw user interface"""
        pass

    def on_channel_volume_changed(
        self, channel_count, new_channel_volume_array, changed_channel, event_context
    ):
        """pycaw user interface"""
        pass

    def on_grouping_param_changed(self, new_grouping_param, event_context):
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
    Helper for audio device volume callbacks.

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


class MMNotificationClient(COMObject):
    """
    Helper for audio endpoint device callbacks.

    Methods
    -------
    Override the following method(s):

    def on_default_device_changed(flow, flow_id, role, role_id, default_device_id):
        Is fired, when the default endpoint device for a role changed.
            flow : str
                String explaining the data-flow direction.
            flow_id: int
                Id of the data-flow direction.
            role : str
                String explaining the role of the device.
            role_id: int
                Id of the role.
            default-device_id: str
                String containing the default device id.

    def on_device_added(self, added_device_id):
        Is fired when a new endpoint device is added.
            added_device_id: str
                String containing the added device id.

    def on_device_removed(self, added_device_id):
        Is fired when a new endpoint device is removed.
            removed_device_id: str
                String containing the removed device id.

    def on_device_state_changed(self, device_id, new_state, new_state_id):
        Is fired when the state of an endpoint device has changed.
            device_id: str
                String containing the id of the device that has changed state.
            new_state: str
                String containing the new state.
            new_state_id: int
                ID of the new state.

    def on_property_value_changed(self, device_id, property_struct, fmtid, pid):
        Is fired when the value of a property belonging to an audio endpoint device
        has changed.
            device_id: str
                String containing the id of the device for which a property is changed.
            property_struct: pycaw.api.mmdeviceapi.depend.structures.PROPERTYKEY
                A structure containing an unique GUID for the property and a PID
                (property identifier).
            fmtid: comtypes.GUID
                GUID of the changed property.
            pid: int
                PID of the changed property.
    """

    _com_interfaces_ = (IMMNotificationClient,)

    DeviceStates = {1: "Active", 2: "Disabled", 4: "NotPresent", 8: "Unplugged"}
    Roles = ["eConsole", "eMultimedia", "eCommunications", "ERole_enum_count"]
    DataFlow = ["eRender", "eCapture", "eAll", "EDataFlow_enum_count"]

    def OnDefaultDeviceChanged(self, flow_id, role_id, default_device_id):
        flow = self.DataFlow[flow_id]
        role = self.Roles[role_id]
        self.on_default_device_changed(flow, flow_id, role, role_id, default_device_id)

    def OnDeviceAdded(self, added_device_id):
        self.on_device_added(added_device_id)

    def OnDeviceRemoved(self, removed_device_id):
        self.on_device_removed(removed_device_id)

    def OnDeviceStateChanged(self, device_id, new_state_id):
        new_state = self.DeviceStates[new_state_id]
        self.on_device_state_changed(device_id, new_state, new_state_id)

    def OnPropertyValueChanged(self, device_id, property_struct):
        fmtid = property_struct.fmtid
        pid = property_struct.pid
        self.on_property_value_changed(device_id, property_struct, fmtid, pid)

    def on_default_device_changed(
        self, flow, flow_id, role, role_id, default_device_id
    ):
        """pycaw user interface"""
        pass

    def on_device_added(self, added_device_id):
        """pycaw user interface"""
        pass

    def on_device_removed(self, removed_device_id):
        """pycaw user interface"""
        pass

    def on_device_state_changed(self, device_id, new_state, new_state_id):
        """pycaw user interface"""
        pass

    def on_property_value_changed(self, device_id, property_struct, fmtid, pid):
        """pycaw user interface"""
        pass
