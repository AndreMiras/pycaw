import warnings
from ctypes import POINTER, cast

import comtypes
import psutil
from _ctypes import COMError

from pycaw.api.audioclient import ISimpleAudioVolume
from pycaw.api.audiopolicy import IAudioSessionControl2, IAudioSessionManager2
from pycaw.api.endpointvolume import IAudioEndpointVolume
from pycaw.api.mmdeviceapi import IMMDeviceEnumerator
from pycaw.constants import (
    DEVICE_STATE,
    STGM,
    AudioDeviceState,
    CLSID_MMDeviceEnumerator,
    EDataFlow,
    ERole,
    IID_Empty,
)


class AudioDevice:
    """
    http://stackoverflow.com/a/20982715/185510
    """

    def __init__(self, id, state, properties, dev):
        self.id = id
        self.state = state
        self.properties = properties
        self._dev = dev
        self._volume = None

    def __str__(self):
        return "AudioDevice: %s" % (self.FriendlyName)

    @property
    def FriendlyName(self):
        DEVPKEY_Device_FriendlyName = (
            "{a45c254e-df1c-4efd-8020-67d146a850e0} 14".upper()
        )
        value = self.properties.get(DEVPKEY_Device_FriendlyName)
        return value

    @property
    def EndpointVolume(self):
        if self._volume is None:
            iface = self._dev.Activate(
                IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None
            )
            self._volume = cast(iface, POINTER(IAudioEndpointVolume))
        return self._volume


class AudioSession:
    """
    http://stackoverflow.com/a/20982715/185510
    """

    def __init__(self, audio_session_control2):
        self._ctl = audio_session_control2
        self._process = None
        self._volume = None
        self._callback = None

    def __str__(self):
        s = self.DisplayName
        if s:
            return "DisplayName: " + s
        if self.Process is not None:
            return "Process: " + self.Process.name()
        return "Pid: %s" % (self.ProcessId)

    @property
    def Process(self):
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
        s = self._ctl.GetDisplayName()
        return s

    @DisplayName.setter
    def DisplayName(self, value):
        s = self._ctl.GetDisplayName()
        if s != value:
            self._ctl.SetDisplayName(value, IID_Empty)

    @property
    def IconPath(self):
        s = self._ctl.GetIconPath()
        return s

    @IconPath.setter
    def IconPath(self, value):
        s = self._ctl.GetIconPath()
        if s != value:
            self._ctl.SetIconPath(value, IID_Empty)

    @property
    def SimpleAudioVolume(self):
        if self._volume is None:
            self._volume = self._ctl.QueryInterface(ISimpleAudioVolume)
        return self._volume

    def register_notification(self, callback):
        if self._callback is None:
            self._callback = callback
            self._ctl.RegisterAudioSessionNotification(self._callback)

    def unregister_notification(self):
        if self._callback:
            self._ctl.UnregisterAudioSessionNotification(self._callback)


class AudioUtilities:
    """
    http://stackoverflow.com/a/20982715/185510
    """

    @staticmethod
    def GetSpeakers():
        """
        get the speakers (1st render + multimedia) device
        """
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        speakers = deviceEnumerator.GetDefaultAudioEndpoint(
            EDataFlow.eRender.value, ERole.eMultimedia.value
        )
        return speakers

    @staticmethod
    def GetMicrophone():
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
    def GetAudioSessionManager():
        speakers = AudioUtilities.GetSpeakers()
        if speakers is None:
            return None
        # win7+ only
        o = speakers.Activate(IAudioSessionManager2._iid_, comtypes.CLSCTX_ALL, None)
        mgr = o.QueryInterface(IAudioSessionManager2)
        return mgr

    @staticmethod
    def GetAllSessions():
        audio_sessions = []
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
    def GetProcessSession(id):
        for session in AudioUtilities.GetAllSessions():
            if session.ProcessId == id:
                return session
            # session.Dispose()
        return None

    @staticmethod
    def CreateDevice(dev):
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
                # TODO
                # PropVariantClear(byref(value))
                name = str(pk)
                properties[name] = v
        audioState = AudioDeviceState(state)
        return AudioDevice(id, audioState, properties, dev)

    @staticmethod
    def GetAllDevices():
        devices = []
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )
        if deviceEnumerator is None:
            return devices

        collection = deviceEnumerator.EnumAudioEndpoints(
            EDataFlow.eAll.value, DEVICE_STATE.MASK_ALL.value
        )
        if collection is None:
            return devices

        count = collection.GetCount()
        for i in range(count):
            dev = collection.Item(i)
            if dev is not None:
                devices.append(AudioUtilities.CreateDevice(dev))
        return devices
