"""
Python wrapper around the Core Audio Windows API.
"""
import psutil
import comtypes
from enum import Enum
from ctypes import cast, HRESULT, POINTER, Structure, Union, c_float, c_void_p
from ctypes.wintypes import BOOL, VARIANT_BOOL, WORD, DWORD, UINT, INT, LONG, \
    ULARGE_INTEGER, LPWSTR
from comtypes import GUID, COMMETHOD
from comtypes.automation import VARTYPE, VT_BOOL, VT_LPWSTR, VT_UI4, VT_CLSID

IID_Empty = GUID(
    '{00000000-0000-0000-0000-000000000000}')
IID_IAudioEndpointVolume = GUID(
    '{5CDF2C82-841E-4546-9722-0CF74078229A}')
IID_IPropertyStore = GUID(
    '{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}')
IID_IMMDevice = GUID(
    '{D666063F-1587-4E43-81F1-B948E807363F}')
IID_IMMDeviceCollection = GUID(
    '{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}')
IID_IMMDeviceEnumerator = GUID(
    '{A95664D2-9614-4F35-A746-DE8DB63617E6}')
IID_IAudioSessionEnumerator = GUID(
    '{E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}')
IID_IAudioSessionManager = GUID(
    '{BFA971F1-4d5e-40bb-935e-967039bfbee4}')
IID_IAudioSessionManager2 = GUID(
    '{77aa99a0-1bd6-484f-8bc7-2c654c9a9b6f}')
IID_IAudioSessionControl = GUID(
    '{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}')
IID_IAudioSessionControl2 = GUID(
    '{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}')

CLSID_MMDeviceEnumerator = GUID(
    '{BCDE0395-E52F-467C-8E3D-C4579291692E}')


class PROPVARIANT_UNION(Union):
        _fields_ = [
            ('lVal', LONG),
            ('uhVal', ULARGE_INTEGER),
            ('boolVal', VARIANT_BOOL),
            ('pwszVal', LPWSTR),
            ('puuid', GUID),
        ]


class PROPVARIANT(Structure):
    _fields_ = [
        ('vt', VARTYPE),
        ('reserved1', WORD),
        ('reserved2', WORD),
        ('reserved3', WORD),
        ('union', PROPVARIANT_UNION),
    ]

    def GetValue(self):
        vt = self.vt
        if vt == VT_BOOL:
            return self.union.boolVal != 0
        elif vt == VT_LPWSTR:
            # return Marshal.PtrToStringUni(union.pwszVal)
            return self.union.pwszVal
        elif vt == VT_UI4:
            return self.union.lVal
        elif vt == VT_CLSID:
            # TODO
            # return (Guid)Marshal.PtrToStructure(union.puuid, typeof(Guid))
            return
        else:
            return unicode(vt) + u":?"


class ERole(Enum):
    eConsole = 0
    eMultimedia = 1
    eCommunications = 2
    ERole_enum_count = 3


class EDataFlow(Enum):
    eRender = 0
    eCapture = 1
    eAll = 2
    EDataFlow_enum_count = 3


class DEVICE_STATE(Enum):
    ACTIVE = 0x00000001
    DISABLED = 0x00000002
    NOTPRESENT = 0x00000004
    UNPLUGGED = 0x00000008
    MASK_ALL = 0x0000000F


class AudioDeviceState(Enum):
    Active = 0x1
    Disabled = 0x2
    NotPresent = 0x4
    Unplugged = 0x8


class STGM(Enum):
    STGM_READ = 0x00000000


class IAudioEndpointVolume(comtypes.IUnknown):
    _iid_ = IID_IAudioEndpointVolume
    _methods_ = (
        # HRESULT RegisterControlChangeNotify(
        # [in] IAudioEndpointVolumeCallback *pNotify);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT UnregisterControlChangeNotify(
        # [in] IAudioEndpointVolumeCallback *pNotify);
        COMMETHOD([], HRESULT, 'NotImpl2'),
        # HRESULT GetChannelCount([out] UINT *pnChannelCount);
        COMMETHOD([], HRESULT, 'GetChannelCount',
                  (['out'], POINTER(UINT), 'pnChannelCount')),
        # HRESULT SetMasterVolumeLevel(
        # [in] float fLevelDB, [in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'SetMasterVolumeLevel',
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT SetMasterVolumeLevelScalar(
        # [in] float fLevel, [in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'SetMasterVolumeLevelScalar',
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT GetMasterVolumeLevel([out] float *pfLevelDB);
        COMMETHOD([], HRESULT, 'GetMasterVolumeLevel',
                  (['out'], POINTER(c_float), 'pfLevelDB')),
        # HRESULT GetMasterVolumeLevelScalar([out] float *pfLevel);
        COMMETHOD([], HRESULT, 'GetMasterVolumeLevelScalar',
                  (['out'], POINTER(c_float), 'pfLevelDB')),
        # HRESULT SetChannelVolumeLevel(
        # [in] UINT nChannel,
        # [in] float fLevelDB,
        # [in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'SetChannelVolumeLevel',
                  (['in'], UINT, 'nChannel'),
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT SetChannelVolumeLevelScalar(
        # [in] UINT nChannel,
        # [in] float fLevel,
        # [in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'SetChannelVolumeLevelScalar',
                  (['in'], DWORD, 'nChannel'),
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT GetChannelVolumeLevel(
        # [in]  UINT nChannel,
        # [out] float *pfLevelDB);
        COMMETHOD([], HRESULT, 'GetChannelVolumeLevel',
                  (['in'], UINT, 'nChannel'),
                  (['out'], POINTER(c_float), 'pfLevelDB')),
        # HRESULT GetChannelVolumeLevelScalar(
        # [in]  UINT nChannel,
        # [out] float *pfLevel);
        COMMETHOD([], HRESULT, 'GetChannelVolumeLevelScalar',
                  (['in'], DWORD, 'nChannel'),
                  (['out'], POINTER(c_float), 'pfLevelDB')),
        # HRESULT SetMute([in] BOOL bMute, [in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'SetMute',
                  (['in'], BOOL, 'bMute'),
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT GetMute([out] BOOL *pbMute);
        COMMETHOD([], HRESULT, 'GetMute',
                  (['out'], POINTER(BOOL), 'pbMute')),
        # HRESULT GetVolumeStepInfo(
        # [out] UINT *pnStep,
        # [out] UINT *pnStepCount);
        COMMETHOD([], HRESULT, 'GetVolumeStepInfo',
                  (['out'], POINTER(c_float), 'pnStep'),
                  (['out'], POINTER(c_float), 'pnStepCount')),
        # HRESULT VolumeStepUp([in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'VolumeStepUp',
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT VolumeStepDown([in] LPCGUID pguidEventContext);
        COMMETHOD([], HRESULT, 'VolumeStepDown',
                  (['in'], POINTER(GUID), 'pguidEventContext')),
        # HRESULT QueryHardwareSupport([out] DWORD *pdwHardwareSupportMask);
        COMMETHOD([], HRESULT, 'QueryHardwareSupport',
                  (['out'], POINTER(DWORD), 'pdwHardwareSupportMask')),
        # HRESULT GetVolumeRange(
        # [out] float *pfLevelMinDB,
        # [out] float *pfLevelMaxDB,
        # [out] float *pfVolumeIncrementDB);
        COMMETHOD([], HRESULT, 'GetVolumeRange',
                  (['out'], POINTER(c_float), 'pfMin'),
                  (['out'], POINTER(c_float), 'pfMax'),
                  (['out'], POINTER(c_float), 'pfIncr')))


class IAudioSessionControl(comtypes.IUnknown):
    _iid_ = IID_IAudioSessionControl
    _methods_ = (
        # HRESULT GetState ([out] AudioSessionState *pRetVal);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT GetDisplayName([out] LPWSTR *pRetVal);
        COMMETHOD([], HRESULT, 'GetDisplayName',
                  (['out'], POINTER(LPWSTR), 'pRetVal')),
        # HRESULT SetDisplayName(
        # [in] LPCWSTR Value,
        # [in] LPCGUID EventContext);
        COMMETHOD([], HRESULT, 'NotImpl2'),
        # HRESULT GetIconPath([out] LPWSTR *pRetVal);
        COMMETHOD([], HRESULT, 'NotImpl3'),
        # HRESULT SetIconPath(
        # [in] LPCWSTR Value,
        # [in] LPCGUID EventContext);
        COMMETHOD([], HRESULT, 'NotImpl4'),
        # HRESULT GetGroupingParam([out] GUID *pRetVal);
        COMMETHOD([], HRESULT, 'NotImpl5'),
        # HRESULT SetGroupingParam(
        # [in] LPCGUID Grouping,
        # [in] LPCGUID EventContext);
        COMMETHOD([], HRESULT, 'NotImpl6'),
        # HRESULT RegisterAudioSessionNotification(
        # [in] IAudioSessionEvents *NewNotifications);
        COMMETHOD([], HRESULT, 'NotImpl7'),
        # HRESULT UnregisterAudioSessionNotification(
        # [in] IAudioSessionEvents *NewNotifications);
        COMMETHOD([], HRESULT, 'NotImpl8'))


class IAudioSessionControl2(IAudioSessionControl):
    _iid_ = IID_IAudioSessionControl2
    _methods_ = (
        # HRESULT GetSessionIdentifier([out] LPWSTR *pRetVal);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT GetSessionInstanceIdentifier([out] LPWSTR *pRetVal);
        COMMETHOD([], HRESULT, 'NotImpl2'),
        # HRESULT GetProcessId([out] DWORD *pRetVal);
        COMMETHOD([], HRESULT, 'GetProcessId',
                  (['out'], POINTER((DWORD)), 'pRetVal')),
        # HRESULT IsSystemSoundsSession();
        COMMETHOD([], HRESULT, 'IsSystemSoundsSession'))


class IAudioSessionEnumerator(comtypes.IUnknown):
    _iid_ = IID_IAudioSessionEnumerator
    _methods_ = (
        # HRESULT GetCount([out] int *SessionCount);
        COMMETHOD([], HRESULT, 'GetCount',
                  (['out'], POINTER(INT), 'SessionCount')),
        # HRESULT GetSession(
        # [in] int SessionCount,
        # [out] IAudioSessionControl **Session);
        COMMETHOD([], HRESULT, 'GetSession',
                  (['in'], INT, 'SessionCount'),
                  (['out'],
                   POINTER(POINTER(IAudioSessionControl)), 'Session')))


class IAudioSessionManager(comtypes.IUnknown):
    _iid_ = IID_IAudioSessionManager
    _methods_ = (
        # HRESULT GetAudioSessionControl(
        # [in] LPCGUID AudioSessionGuid,
        # [in] DWORD StreamFlags,
        # [out] IAudioSessionControl **SessionControl);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT GetSimpleAudioVolume(
        # [in] LPCGUID AudioSessionGuid,
        # [in] DWORD CrossProcessSession,
        # [out] ISimpleAudioVolume **AudioVolume);
        COMMETHOD([], HRESULT, 'NotImpl2'))


class IAudioSessionManager2(IAudioSessionManager):
    _iid_ = IID_IAudioSessionManager2
    _methods_ = (
        # HRESULT GetSessionEnumerator(
        # [out] IAudioSessionEnumerator **SessionList);
        COMMETHOD([], HRESULT, 'GetSessionEnumerator',
                  (['out'],
                  POINTER(POINTER(IAudioSessionEnumerator)), 'SessionList')),
        # HRESULT RegisterSessionNotification(
        # IAudioSessionNotification *SessionNotification);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT UnregisterSessionNotification(
        # IAudioSessionNotification *SessionNotification);
        COMMETHOD([], HRESULT, 'NotImpl2'),
        # HRESULT RegisterDuckNotification(
        # LPCWSTR SessionID,
        # IAudioVolumeDuckNotification *duckNotification);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT UnregisterDuckNotification(
        # IAudioVolumeDuckNotification *duckNotification);
        COMMETHOD([], HRESULT, 'NotImpl2'))


class PROPERTYKEY(Structure):
    _fields_ = [
        ('fmtid', GUID),
        ('pid', DWORD),
    ]

    def __unicode__(self):
        return unicode(self.fmtid) + u" " + unicode(self.pid)

    def __str__(self):
        return unicode(self).encode('utf-8')


class IPropertyStore(comtypes.IUnknown):
    _iid_ = IID_IPropertyStore
    _methods_ = (
        # HRESULT GetCount([out] DWORD *cProps);
        COMMETHOD([], HRESULT, 'GetCount',
                  (['out'], POINTER(DWORD), 'cProps')),
        # HRESULT GetAt(
        # [in] DWORD iProp,
        # [out] PROPERTYKEY *pkey);
        COMMETHOD([], HRESULT, 'GetAt',
                  (['in'], DWORD, 'iProp'),
                  (['out'], POINTER(PROPERTYKEY), 'pkey')),
        # HRESULT GetValue(
        # [in] REFPROPERTYKEY key,
        # [out] PROPVARIANT *pv);
        COMMETHOD([], HRESULT, 'GetValue',
                  (['in'], POINTER(PROPERTYKEY), 'key'),
                  (['out'], POINTER(PROPVARIANT), 'pv')),
        # HRESULT SetValue([out] LPWSTR *ppstrId);
        COMMETHOD([], HRESULT, 'SetValue',
                  (['out'], POINTER(LPWSTR), 'ppstrId')),
        # HRESULT Commit();
        COMMETHOD([], HRESULT, 'Commit'))


class IMMDevice(comtypes.IUnknown):
    _iid_ = IID_IMMDevice
    _methods_ = (
        # HRESULT Activate(
        # [in] REFIID iid,
        # [in] DWORD dwClsCtx,
        # [in] PROPVARIANT *pActivationParams,
        # [out] void **ppInterface);
        COMMETHOD([], HRESULT, 'Activate',
                  (['in'], POINTER(GUID), 'iid'),
                  (['in'], DWORD, 'dwClsCtx'),
                  (['in'], POINTER(PROPVARIANT), 'pActivationParams'),
                  (['out'], POINTER(c_void_p), 'ppInterface')),
        # HRESULT OpenPropertyStore(
        # [in] DWORD stgmAccess,
        # [out] IPropertyStore **ppProperties);
        COMMETHOD([], HRESULT, 'OpenPropertyStore',
                  (['in'], DWORD, 'stgmAccess'),
                  (['out'],
                  POINTER(POINTER(IPropertyStore)), 'ppProperties')),
        # HRESULT GetId([out] LPWSTR *ppstrId);
        COMMETHOD([], HRESULT, 'GetId',
                  (['out'], POINTER(LPWSTR), 'ppstrId')),
        # HRESULT GetState([out] DWORD *pdwState);
        COMMETHOD([], HRESULT, 'GetState',
                  (['out'], POINTER(DWORD), 'pdwState')))


class IMMDeviceCollection(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceCollection
    _methods_ = (
        # HRESULT GetCount([out] UINT *pcDevices);
        COMMETHOD([], HRESULT, 'GetCount',
                  (['out'], POINTER(UINT), 'pcDevices')),
        # HRESULT Item([in] UINT nDevice, [out] IMMDevice **ppDevice);
        COMMETHOD([], HRESULT, 'Item',
                  (['in'], UINT, 'nDevice'),
                  (['out'], POINTER(POINTER(IMMDevice)), 'ppDevice')))


class IMMDeviceEnumerator(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceEnumerator
    _methods_ = (
        # HRESULT EnumAudioEndpoints(
        # [in] EDataFlow dataFlow,
        # [in] DWORD dwStateMask,
        # [out] IMMDeviceCollection **ppDevices);
        COMMETHOD([], HRESULT, 'EnumAudioEndpoints',
                  (['in'], DWORD, 'dataFlow'),
                  (['in'], DWORD, 'dwStateMask'),
                  (['out'],
                  POINTER(POINTER(IMMDeviceCollection)), 'ppDevices')),
        # HRESULT GetDefaultAudioEndpoint(
        # [in] EDataFlow dataFlow,
        # [in] ERole role,
        # [out] IMMDevice **ppDevice);
        COMMETHOD([], HRESULT, 'GetDefaultAudioEndpoint',
                  # (['in'], EDataFlow, 'dataFlow'),
                  (['in'], DWORD, 'dataFlow'),
                  # (['in'], ERole, 'role'),
                  (['in'], DWORD, 'role'),
                  (['out'], POINTER(POINTER(IMMDevice)), 'ppDevices')))


class AudioDevice(object):
    """
    http://stackoverflow.com/a/20982715/185510
    """
    def __init__(self, id, state, properties):
        self.id = id
        self.state = state
        self.properties = properties

    def __str__(self):
        return unicode(self).encode('utf-8')

    def __unicode__(self):
        return u"AudioDevice: " + unicode(self.FriendlyName)

    @property
    def FriendlyName(self):
        DEVPKEY_Device_FriendlyName = \
            u"{a45c254e-df1c-4efd-8020-67d146a850e0} 14".upper()
        value = self.properties.get(DEVPKEY_Device_FriendlyName)
        return value


class AudioSession(object):
    """
    http://stackoverflow.com/a/20982715/185510
    """

    def __init__(self, audio_session_control2):
        self._ctl = audio_session_control2
        self._process = None

    def __enter__(self):
        pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._ctl is not None:
            # Marshal.ReleaseComObject(_ctl);
            self._ctl = None

    def __str__(self):
        return unicode(self).encode('utf-8')

    def __unicode__(self):
        s = self.DisplayName
        if s:
            return "DisplayName: " + s
        if self.Process is not None:
            return "Process: " + self.Process.name()
        return "Pid: " + self.ProcessId

    @property
    def Process(self):
        if self._process is None and self.ProcessId != 0:
            self._process = psutil.Process(self.ProcessId)
        return self._process

    @property
    def ProcessId(self):
        self.CheckDisposed()
        i = self._ctl.GetProcessId()
        return i

    @property
    def Identifier(self):
        self.CheckDisposed()
        s = self._ctl.GetSessionIdentifier()
        return s

    @property
    def InstanceIdentifier(self):
        self.CheckDisposed()
        s = self._ctl.GetSessionInstanceIdentifier()
        return s

    @property
    def State(self):
        self.CheckDisposed()
        s = self._ctl.GetState()
        return s

    @property
    def GroupingParam(self):
        self.CheckDisposed()
        g = self._ctl.GetGroupingParam()
        return g

    @GroupingParam.setter
    def GroupingParam(self, value):
        self.CheckDisposed()
        self._ctl.SetGroupingParam(value, IID_Empty)

    @property
    def DisplayName(self):
        self.CheckDisposed()
        s = self._ctl.GetDisplayName()
        return s

    @DisplayName.setter
    def DisplayName(self, value):
        self.CheckDisposed()
        s = self._ctl.GetDisplayName()
        if s != value:
            self._ctl.SetDisplayName(value, IID_Empty)

    @property
    def IconPath(self):
        self.CheckDisposed()
        s = self._ctl.GetIconPath()
        return s

    @IconPath.setter
    def IconPath(self, value):
        self.CheckDisposed()
        s = self._ctl.GetIconPath()
        if s != value:
            self._ctl.SetIconPath(value, IID_Empty)

    def CheckDisposed(self):
        if self._ctl is None:
            raise Exception("ObjectDisposedException()")


class AudioUtilities(object):
    """
    http://stackoverflow.com/a/20982715/185510
    """
    @staticmethod
    def GetSpeakers():
        """
        get the speakers (1st render + multimedia) device
        """
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER)
        speakers = deviceEnumerator.GetDefaultAudioEndpoint(
                    EDataFlow.eRender.value, ERole.eMultimedia.value)
        return speakers

    @staticmethod
    def GetAudioSessionManager():
        speakers = AudioUtilities.GetSpeakers()
        if speakers is None:
            return None
        # win7+ only
        o = speakers.Activate(
            IAudioSessionManager2._iid_, comtypes.CLSCTX_ALL, None)
        mgr = cast(o, POINTER(IAudioSessionManager2))
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
            ctl2 = cast(ctl, POINTER(IAudioSessionControl2))
            if ctl2 is not None:
                audio_sessions.append(AudioSession(ctl2))
        # Marshal.ReleaseComObject(sessionEnumerator);
        # Marshal.ReleaseComObject(mgr);
        return audio_sessions

    @staticmethod
    def GetProcessSession(id):
        for session in AudioUtilities.GetAllSessions():
            if session.ProcessId == id:
                return session
            # session.Dispose()
        return None

