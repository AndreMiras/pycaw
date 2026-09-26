import warnings
from _ctypes import COMError
from collections.abc import Iterable
from typing import Any, Literal, overload

import comtypes
import psutil

from pycaw.api.audioclient import IChannelAudioVolume, ISimpleAudioVolume
from pycaw.api.audiopolicy import IAudioSessionControl2, IAudioSessionManager2
from pycaw.api.endpointvolume import IAudioEndpointVolume
from pycaw.api.mmdeviceapi import IMMDeviceEnumerator, IMMEndpoint
from pycaw.api.policyconfig import IPolicyConfig
from pycaw.constants import (
    DEVICE_STATE,
    STGM,
    AudioDeviceState,
    CLSID_CPolicyConfigClient,
    CLSID_MMDeviceEnumerator,
    EDataFlow,
    ERole,
    IID_Empty,
)


class AudioDevice:
    """
    https://stackoverflow.com/a/20982715/185510
    """

    def __init__(
        self,
        id: str,
        state: AudioDeviceState,
        properties: dict[str, Any],
        dev: Any,
    ) -> None:
        self.id = id
        self.state = state
        self.properties = properties
        self._dev = dev
        self._volume: Any = None
        self._audio_session_manager: Any = None

    def __str__(self) -> str:
        return "AudioDevice: %s" % (self.FriendlyName)

    @property
    def FriendlyName(self) -> str | None:
        DEVPKEY_Device_FriendlyName = (
            "{a45c254e-df1c-4efd-8020-67d146a850e0} 14".upper()
        )
        value = self.properties.get(DEVPKEY_Device_FriendlyName)

        # Fallback to DeviceDesc if FriendlyName is unavailable
        if value is None or (isinstance(value, str) and not value):
            DEVPKEY_Device_DeviceDesc = (
                "{a45c254e-df1c-4efd-8020-67d146a850e0} 2".upper()
            )
            value = self.properties.get(DEVPKEY_Device_DeviceDesc)

        return value

    @property
    def EndpointVolume(self) -> Any:
        if self._volume is None:
            iface = self._dev.Activate(
                IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None
            )
            self._volume = iface.QueryInterface(IAudioEndpointVolume)
        return self._volume

    @property
    def volume_percent(self) -> float:
        """
        Master volume of this device as a percentage (0.0-100.0), the same
        scale the Windows volume mixer shows.

        Convenience wrapper around `GetMasterVolumeLevelScalar()` /
        `SetMasterVolumeLevelScalar()`, so users don't reach for the decibel
        based `SetMasterVolumeLevel()` by mistake. The setter clamps the
        value to the 0-100 range.
        """
        return self.EndpointVolume.GetMasterVolumeLevelScalar() * 100

    @volume_percent.setter
    def volume_percent(self, percent: float) -> None:
        percent = max(0.0, min(100.0, percent))
        self.EndpointVolume.SetMasterVolumeLevelScalar(percent / 100, None)

    @property
    def AudioSessionManager(self) -> Any:
        if self._audio_session_manager is None:
            # win7+ only
            iface = self._dev.Activate(
                IAudioSessionManager2._iid_, comtypes.CLSCTX_ALL, None
            )
            self._audio_session_manager = iface.QueryInterface(IAudioSessionManager2)
        return self._audio_session_manager


class AudioSession:
    """
    https://stackoverflow.com/a/20982715/185510
    """

    def __init__(self, audio_session_control2: Any) -> None:
        self._ctl = audio_session_control2
        self._process: psutil.Process | None = None
        self._volume: Any = None
        self._channelVolume: Any = None
        self._callback: Any = None

    def __str__(self) -> str:
        s = self.DisplayName
        if s:
            return "DisplayName: " + s
        if self.Process is not None:
            return "Process: " + self.Process.name()
        return "Pid: %s" % (self.ProcessId)

    @property
    def Process(self) -> psutil.Process | None:
        """Return the session's process, or ``None`` when it has no process.

        The special Windows System Sounds session has process ID 0, so it does
        not have an associated :class:`psutil.Process`.
        """
        if self._process is None and self.ProcessId != 0:
            try:
                self._process = psutil.Process(self.ProcessId)
            except psutil.NoSuchProcess:
                # for some reason GetProcessId returned an non existing pid
                return None
        return self._process

    @property
    def ProcessId(self) -> int:
        return self._ctl.GetProcessId()

    @property
    def Identifier(self) -> str:
        s = self._ctl.GetSessionIdentifier()
        return s

    @property
    def InstanceIdentifier(self) -> str:
        s = self._ctl.GetSessionInstanceIdentifier()
        return s

    @property
    def State(self) -> int:
        s = self._ctl.GetState()
        return s

    @property
    def GroupingParam(self) -> Any:
        g = self._ctl.GetGroupingParam()
        return g

    @GroupingParam.setter
    def GroupingParam(self, value: Any) -> None:
        self._ctl.SetGroupingParam(value, IID_Empty)

    @property
    def DisplayName(self) -> str:
        """
        Please, note that this returns an empty string if
        the client hadn't called the setter method before.
        """
        s = self._ctl.GetDisplayName()
        return s

    @DisplayName.setter
    def DisplayName(self, value: str) -> None:
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
    def IconPath(self, value: str) -> None:
        s = self._ctl.GetIconPath()
        if s != value:
            self._ctl.SetIconPath(value, IID_Empty)

    @property
    def SimpleAudioVolume(self) -> Any:
        if self._volume is None:
            self._volume = self._ctl.QueryInterface(ISimpleAudioVolume)
        return self._volume

    def channelAudioVolume(self) -> Any:
        if self._channelVolume is None:
            self._channelVolume = self._ctl.QueryInterface(IChannelAudioVolume)
        return self._channelVolume

    def register_notification(self, callback: Any) -> None:
        if self._callback is None:
            self._callback = callback
            self._ctl.RegisterAudioSessionNotification(self._callback)

    def unregister_notification(self) -> None:
        if self._callback is not None:
            self._ctl.UnregisterAudioSessionNotification(self._callback)
            self._callback = None


class AudioUtilities:
    """
    https://stackoverflow.com/a/20982715/185510
    """

    @staticmethod
    def GetSpeakers() -> AudioDevice | None:
        """
        get the speakers (1st render + multimedia) device
        """
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        speakers = deviceEnumerator.GetDefaultAudioEndpoint(
            EDataFlow.eRender.value, ERole.eMultimedia.value
        )
        return AudioUtilities.CreateDevice(speakers)

    @staticmethod
    def GetMicrophone() -> Any:
        """
        get the microphone (1st capture + multimedia) device
        """
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        microphone = deviceEnumerator.GetDefaultAudioEndpoint(
            EDataFlow.eCapture.value, ERole.eMultimedia.value
        )
        return microphone

    @staticmethod
    def GetAudioSessionManager() -> Any:
        speakers = AudioUtilities.GetSpeakers()
        if speakers is None:
            return None
        return speakers.AudioSessionManager

    @staticmethod
    def GetAllSessions() -> list[AudioSession]:
        audio_sessions: list[AudioSession] = []
        mgr = AudioUtilities.GetAudioSessionManager()
        if mgr is None:
            return audio_sessions
        sessionEnumerator = mgr.GetSessionEnumerator()
        count = sessionEnumerator.GetCount()
        for i in range(count):
            ctl = sessionEnumerator.GetSession(i)
            if ctl is None:
                continue
            ctl2 = ctl.QueryInterface(IAudioSessionControl2)
            if ctl2 is not None:
                audio_session = AudioSession(ctl2)
                audio_sessions.append(audio_session)
        return audio_sessions

    @staticmethod
    def GetProcessSession(id: int) -> AudioSession | None:
        for session in AudioUtilities.GetAllSessions():
            if session.ProcessId == id:
                return session
            # session.Dispose()
        return None

    @staticmethod
    def CreateDevice(dev: Any | None) -> AudioDevice | None:
        if dev is None:
            return None
        id = dev.GetId()
        state = dev.GetState()
        properties = {}
        store = dev.OpenPropertyStore(STGM.STGM_READ.value)
        if store is not None:
            propCount = store.GetCount()
            for j in range(propCount):
                try:
                    pk = store.GetAt(j)
                    value = store.GetValue(pk)
                    v = value.GetValue()
                except COMError as exc:
                    warnings.warn(
                        "COMError attempting to get property %r "
                        "from device %r: %r" % (j, dev, exc)
                    )
                    continue
                value.clear()
                name = str(pk)
                properties[name] = v
        audioState = AudioDeviceState(state)
        return AudioDevice(id, audioState, properties, dev)

    @staticmethod
    def GetAllDevices(
        data_flow: int = EDataFlow.eAll.value,
        device_state: int = DEVICE_STATE.MASK_ALL.value,
    ) -> list[AudioDevice]:
        devices: list[AudioDevice] = []
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        if deviceEnumerator is None:
            return devices

        collection = deviceEnumerator.EnumAudioEndpoints(data_flow, device_state)
        if collection is None:
            return devices

        count = collection.GetCount()
        for i in range(count):
            dev = collection.Item(i)
            if dev is not None:
                device = AudioUtilities.CreateDevice(dev)
                if device is not None:
                    devices.append(device)
        return devices

    @staticmethod
    def GetDeviceEnumerator() -> Any:
        """
        Get an instance of IMMDeviceEnumerator.
        """
        device_enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        return device_enumerator

    @staticmethod
    @overload
    def GetEndpointDataFlow(devId: str, outputType: Literal[0] = 0) -> str: ...

    @staticmethod
    @overload
    def GetEndpointDataFlow(devId: str, outputType: Literal[1]) -> int: ...

    @staticmethod
    @overload
    def GetEndpointDataFlow(devId: str, outputType: int) -> str | int: ...

    @staticmethod
    def GetEndpointDataFlow(devId: str, outputType: int = 0) -> str | int:
        """
        Get data flow information of a given endpoint.

        Parameters
        ----------
        devId : str
            ID of the device to query
        outputType : int, optional
            Output format: 0 (default) returns text representation,
            1 returns numeric code

        Returns
        -------
        str or int
            Data flow direction. If outputType=0, returns one of:
            "eRender", "eCapture", "eAll", "EDataFlow_enum_count".
            If outputType=1, returns the numeric value (0-3).
        """
        DataFlow = ["eRender", "eCapture", "eAll", "EDataFlow_enum_count"]
        devEnum = AudioUtilities.GetDeviceEnumerator()
        dev = devEnum.GetDevice(devId)
        value = dev.QueryInterface(IMMEndpoint).GetDataFlow()
        if outputType:
            return value
        else:
            return DataFlow[value]

    @staticmethod
    def SetDefaultDevice(devId: str, roles: Iterable[ERole] | None = None) -> None:
        if roles is None:
            roles = [ERole.eConsole]

        # Try newer IPolicyConfig first, fall back to Vista interface if unavailable
        policy_config: Any = None
        try:
            policy_config = comtypes.CoCreateInstance(
                CLSID_CPolicyConfigClient, IPolicyConfig, comtypes.CLSCTX_ALL
            )
        except (OSError, comtypes.COMError):
            # Windows Vista/7 may only have IPolicyConfigVista
            try:
                from pycaw.api.policyconfig import IPolicyConfigVista

                policy_config = comtypes.CoCreateInstance(
                    CLSID_CPolicyConfigClient,
                    IPolicyConfigVista,
                    comtypes.CLSCTX_ALL,
                )
            except (OSError, comtypes.COMError) as e:
                raise OSError(
                    f"Failed to create PolicyConfig interface. "
                    f"This feature requires Windows Vista or later. "
                    f"Original error: {e}"
                )

        for role in roles:
            hr = policy_config.SetDefaultEndpoint(devId, role.value)
            if hr != 0:
                raise OSError(
                    f"SetDefaultEndpoint failed for role {role} with HRESULT {hr:#x}"
                )
