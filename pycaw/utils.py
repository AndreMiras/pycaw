from __future__ import annotations

import warnings
from typing import List, TypeVar

import comtypes
import psutil
from _ctypes import COMError
from comtypes import IUnknown

from pycaw.api.audioclient import IChannelAudioVolume, ISimpleAudioVolume
from pycaw.api.audiopolicy import IAudioSessionControl2, IAudioSessionManager2
from pycaw.api.endpointvolume import IAudioEndpointVolume
from pycaw.api.mmdeviceapi import IMMDeviceEnumerator, IMMEndpoint
from pycaw.constants import (
    DEVICE_STATE,
    STGM,
    AudioDeviceState,
    CLSID_MMDeviceEnumerator,
    EDataFlow,
    ERole,
    IID_Empty,
)

# Define a type variable that extends IUnknown
COMInterface = TypeVar("COMInterface", bound="IUnknown")


class AudioDevice:
    """
    https://stackoverflow.com/a/20982715/185510
    """

    def __init__(self, dev):
        self._dev = dev
        self._id = None
        self._state = None
        self._properties = None
        self._interfaces: dict[str, COMInterface] = {}

    def __str__(self):
        return "AudioDevice: %s" % (self.FriendlyName)

    def ActivateInterface(self, interface: COMInterface) -> COMInterface:
        interface_id = interface._iid_
        if interface_id is None:
            return None

        interface_name = str(interface_id)

        if interface_name not in self._interfaces:
            activated = self._dev.Activate(interface_id, comtypes.CLSCTX_ALL, None)
            interface = activated.QueryInterface(interface)
            self._interfaces[interface_name] = interface

        return self._interfaces[interface_name]

    @property
    def id(self):
        if self._id is None:
            self._id = self._dev.GetId()
        return self._id

    @property
    def state(self) -> AudioDeviceState:
        if self._state is None:
            state = self._dev.GetState()
            self._state = AudioDeviceState(state)
        return self._state

    @staticmethod
    def getProperty(self, key):
        store = self._dev.OpenPropertyStore(STGM.STGM_READ.value)
        if store is None:
            return None

        propCount = store.GetCount()
        for j in range(propCount):
            pk = store.GetAt(j)
            name = str(pk)

            if name != key:
                continue

            try:
                value = store.GetValue(pk)
                v = value.GetValue()
                value.clear()
                return v
            except COMError as exc:
                warnings.warn(
                    "COMError attempting to get property %r "
                    "from device %r: %r" % (j, id, exc)
                )
                return {}

    @property
    def properties(self) -> dict:
        if self._properties is None:
            store = self._dev.OpenPropertyStore(STGM.STGM_READ.value)
            if store is None:
                return {}

            properties = {}
            propCount = store.GetCount()
            for j in range(propCount):
                pk = store.GetAt(j)

                value = AudioDevice.getProperty(self, pk)
                if value is None:
                    continue

                name = str(pk)
                properties[name] = value
            self._properties = properties
        return self._properties

    @property
    def FriendlyName(self):
        DEVPKEY_Device_FriendlyName = (
            "{a45c254e-df1c-4efd-8020-67d146a850e0} 14".upper()
        )
        value = self.properties.get(DEVPKEY_Device_FriendlyName)
        return value

    @property
    def EndpointVolume(self) -> IAudioEndpointVolume:
        interface = self.ActivateInterface(IAudioEndpointVolume)
        return interface

    @property
    def IsMuted(self) -> bool:
        endpointVolume = self.EndpointVolume
        return endpointVolume.GetMute() == 1

    def Mute(self):
        endpointVolume = self.EndpointVolume
        endpointVolume.SetMute(1, None)

    def UnMute(self):
        endpointVolume = self.EndpointVolume
        endpointVolume.SetMute(0, None)

    @property
    def AudioSessionManager(self) -> IAudioSessionManager2:
        return self.ActivateInterface(IAudioSessionManager2)

    @property
    def Sessions(self) -> List[AudioSession]:
        # Only Active devices should have audio sessions
        if self.state != AudioDeviceState.Active:
            return []

        mgr = self.AudioSessionManager

        # Assume no sessions if there's no session enumerator
        if not hasattr(mgr, "GetSessionEnumerator"):
            return []

        sessions = []

        sessionEnumerator = mgr.GetSessionEnumerator()
        count = sessionEnumerator.GetCount()
        for i in range(count):
            ctl = sessionEnumerator.GetSession(i)
            if ctl is None:
                continue
            ctl2 = ctl.QueryInterface(IAudioSessionControl2)
            if ctl2 is not None:
                audio_session = AudioSession(ctl2)
                sessions.append(audio_session)

        return sessions


class AudioSession:
    """
    https://stackoverflow.com/a/20982715/185510
    """

    def __init__(self, audio_session_control2: IAudioSessionControl2):
        self._ctl = audio_session_control2
        self._process = None
        self._interfaces: dict[str, COMInterface] = {}
        self._callback = None

    def __str__(self):
        s = self.DisplayName
        if s:
            return "DisplayName: " + s
        if self.Process is not None:
            return "Process: " + self.Process.name()
        return "Pid: %s" % (self.ProcessId)

    @property
    def Process(self) -> psutil.Process | None:
        if self._process is None and self.ProcessId != 0:
            try:
                self._process = psutil.Process(self.ProcessId)
            except psutil.NoSuchProcess:
                # for some reason GetProcessId returned an non existing pid
                return None
        return self._process

    @property
    def ProcessId(self):
        return self._ctl.GetProcessId()

    @property
    def Identifier(self):
        s = self._ctl.GetSessionIdentifier()
        return s

    @property
    def InstanceIdentifier(self):
        s = self._ctl.GetSessionInstanceIdentifier()
        return s

    @property
    def State(self):
        s = self._ctl.GetState()
        return s

    @property
    def GroupingParam(self):
        g = self._ctl.GetGroupingParam()
        return g

    @GroupingParam.setter
    def GroupingParam(self, value):
        self._ctl.SetGroupingParam(value, IID_Empty)

    @property
    def DisplayName(self):
        """
        Please, note that this returns an empty string if
        the client hadn't called the setter method before.
        """
        s = self._ctl.GetDisplayName()
        return s

    @DisplayName.setter
    def DisplayName(self, value: str) -> str:
        s = self._ctl.GetDisplayName()
        if s != value:
            self._ctl.SetDisplayName(value, IID_Empty)

    @property
    def IconPath(self) -> str:
        """
        Please, note that this returns an empty string if
        the client hadn't called the setter method before.
        """
        s = self._ctl.GetIconPath()
        return s

    @IconPath.setter
    def IconPath(self, value: str):
        s = self._ctl.GetIconPath()
        if s != value:
            self._ctl.SetIconPath(value, IID_Empty)

    def QueryInterface(self, interface: COMInterface) -> COMInterface:
        interface_name = str(interface._iid_)

        if interface_name not in self._interfaces:
            self._interfaces[interface_name] = self._ctl.QueryInterface(interface)

        return self._interfaces[interface_name]

    @property
    def SimpleAudioVolume(self):
        return self.QueryInterface(ISimpleAudioVolume)

    def channelAudioVolume(self):
        return self.QueryInterface(IChannelAudioVolume)

    def register_notification(self, callback):
        if self._callback is None:
            self._callback = callback
            self._ctl.RegisterAudioSessionNotification(self._callback)

    def unregister_notification(self):
        if self._callback:
            self._ctl.UnregisterAudioSessionNotification(self._callback)


class AudioUtilities:
    """
    https://stackoverflow.com/a/20982715/185510
    """

    @staticmethod
    def GetDefaultEndpoint(dataFlow: EDataFlow = EDataFlow.eAll.value):
        deviceEnumerator = AudioUtilities.GetDeviceEnumerator()
        if deviceEnumerator is None:
            return None

        return deviceEnumerator.GetDefaultAudioEndpoint(
            dataFlow, ERole.eMultimedia.value
        )

    @staticmethod
    def GetSpeakers():
        """
        get the speakers (1st render + multimedia) device
        """
        return AudioUtilities.GetDefaultEndpoint(EDataFlow.eRender.value)

    @staticmethod
    def GetMicrophone():
        """
        get the microphone (1st capture + multimedia) device
        """
        return AudioUtilities.GetDefaultEndpoint(EDataFlow.eCapture.value)

    @staticmethod
    def GetAudioSessionManager() -> IAudioSessionManager2 | None:
        speakers = AudioUtilities.GetSpeakers()
        if speakers is None:
            return None

        device = AudioUtilities.CreateDevice(speakers)
        return device.AudioSessionManager

    @staticmethod
    def GetAllSessions() -> List[AudioSession]:
        # TODO: the current behavior here isn't really getting "all sessions"
        #       Leaving behavior as-is for backward compatibility
        speakers = AudioUtilities.GetSpeakers()
        if speakers is None:
            return []

        device = AudioUtilities.CreateDevice(speakers)
        return device.Sessions

    @staticmethod
    def GetSessions(
        dataFlow: EDataFlow = EDataFlow.eAll.value,
        sessionState: AudioDeviceState = None,
    ) -> List[AudioSession]:
        # Only active devices can have sessions associated with them
        devices = AudioUtilities.GetDevices(dataFlow, DEVICE_STATE.ACTIVE.value)
        if devices is None:
            return []

        sessions = []
        for device in devices:
            deviceSessions = device.Sessions
            if deviceSessions is None:
                continue

            for deviceSession in deviceSessions:
                if sessionState is None or sessionState == deviceSession.State:
                    sessions.append(deviceSession)

        return sessions

    @staticmethod
    def GetPlaybackSessions(
        sessionState: AudioDeviceState = None,
    ) -> List[AudioSession]:
        return AudioUtilities.GetSessions(EDataFlow.eRender.value, sessionState)

    @staticmethod
    def GetRecordingSessions(
        sessionState: AudioDeviceState = None,
    ) -> List[AudioSession]:
        return AudioUtilities.GetSessions(EDataFlow.eCapture.value, sessionState)

    @staticmethod
    def GetProcessSession(id) -> AudioSession | None:
        # Need to look at all sessions to ensure
        # we find the one we're looking for
        for session in AudioUtilities.GetSessions():
            if session is None:
                continue
            if session.ProcessId == id:
                return session
            # session.Dispose()
        return None

    @staticmethod
    def CreateDevice(dev) -> AudioDevice | None:
        if dev is None:
            return None

        device = AudioDevice(dev)

        return device

    @staticmethod
    def CreateDevices(collection) -> List[AudioDevice]:
        if collection is None:
            return []

        devices = []

        count = collection.GetCount()
        for i in range(count):
            dev = collection.Item(i)
            device = AudioUtilities.CreateDevice(dev)

            if device is not None:
                devices.append(device)

        return devices

    @staticmethod
    def GetDevices(
        flow=EDataFlow.eAll.value, deviceState=DEVICE_STATE.ACTIVE.value
    ) -> List[AudioDevice]:
        """
        Get devices based on filteres for flow direction and device state.
        Default to returning active devices.
        """
        deviceEnumerator = AudioUtilities.GetDeviceEnumerator()
        if deviceEnumerator is None:
            return []

        collection = deviceEnumerator.EnumAudioEndpoints(flow, deviceState)
        return AudioUtilities.CreateDevices(collection)

    @staticmethod
    def GetAllDevices() -> List[IMMDeviceEnumerator]:
        deviceEnumerator = AudioUtilities.GetDeviceEnumerator()
        if deviceEnumerator is None:
            return []

        collection = deviceEnumerator.EnumAudioEndpoints(
            EDataFlow.eAll.value, DEVICE_STATE.MASK_ALL.value
        )
        return AudioUtilities.CreateDevices(collection)

    @staticmethod
    def GetDeviceEnumerator() -> IMMDeviceEnumerator:
        """
        Get an instance of IMMDeviceEnumerator.
        """
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        return deviceEnumerator

    @staticmethod
    def GetDevice(devId):
        """
        Get AudioDevice.
        One input argument:
            - devId: id of the device
        """
        deviceEnumerator = AudioUtilities.GetDeviceEnumerator()
        dev = deviceEnumerator.GetDevice(devId)
        return AudioUtilities.CreateDevice(dev)

    @staticmethod
    def GetEndpointDataFlow(devId, outputType=0):
        """
        Get data flow information of a given endpoint.
        Two input arguments:
            - devId: id of the device
            - outputType: 0 (default) for text, 1 for code.
        """
        DataFlow = ["eRender", "eCapture", "eAll", "EDataFlow_enum_count"]
        deviceEnumerator = AudioUtilities.GetDeviceEnumerator()
        dev = deviceEnumerator.GetDevice(devId)
        value = dev.QueryInterface(IMMEndpoint).GetDataFlow()
        if outputType:
            return value
        else:
            return DataFlow[value]
