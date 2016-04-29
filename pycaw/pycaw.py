"""
Python wrapper around the Core Audio Windows API.
"""
import psutil
import comtypes
from enum import Enum
from ctypes import cast, HRESULT, POINTER, Structure, c_float, c_void_p
from ctypes.wintypes import BOOL, DWORD, UINT, INT, LPWSTR
from comtypes import GUID, COMMETHOD

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


class PROPVARIANT(DWORD):
    pass


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


class IMMDeviceCollection(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceCollection
    _methods_ = (
        # HRESULT GetCount([out] UINT *pcDevices);
        COMMETHOD([], HRESULT, 'NotImpl1'),
        # HRESULT Item([in] UINT nDevice, [out] IMMDevice **ppDevice);
        COMMETHOD([], HRESULT, 'NotImpl2'))


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


class IPropertyStore(comtypes.IUnknown):
    _iid_ = IID_IPropertyStore
    _methods_ = (
        # HRESULT GetCount([out] DWORD *cProps);
        COMMETHOD([], HRESULT, 'GetAt',
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
                  # (['in'], REFPROPERTYKEY, 'key'),
                  (['in'], POINTER(PROPERTYKEY), 'key'),
                  (['out'], POINTER(PROPVARIANT), 'pv')),
        # HRESULT SetValue([out] LPWSTR *ppstrId);
        COMMETHOD([], HRESULT, 'SetValue',
                  (['out'], POINTER(LPWSTR), 'ppstrId')),
        # HRESULT Commit([out] DWORD *pdwState);
        COMMETHOD([], HRESULT, 'Commit',
                  (['out'], POINTER(DWORD), 'pdwState')))


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
        COMMETHOD([], HRESULT, 'NotImpl2'),
        # HRESULT GetState([out] DWORD *pdwState);
        COMMETHOD([], HRESULT, 'NotImpl3'))


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
                  (['in'], DWORD, 'dataFlow'),
                  (['in'], DWORD, 'role'),
                  (['out'], POINTER(POINTER(IMMDevice)), 'ppDevices')))

