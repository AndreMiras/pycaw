from ctypes import HRESULT, POINTER
from ctypes.wintypes import DWORD, LPCWSTR, LPWSTR, UINT

from comtypes import COMMETHOD, GUID, IUnknown

from .depend import IPropertyStore


class IMMDevice(IUnknown):
    _iid_ = GUID("{D666063F-1587-4E43-81F1-B948E807363F}")
    _methods_ = (
        # HRESULT Activate(
        # [in] REFIID iid,
        # [in] DWORD dwClsCtx,
        # [in] PROPVARIANT *pActivationParams,
        # [out] void **ppInterface);
        COMMETHOD(
            [],
            HRESULT,
            "Activate",
            (["in"], POINTER(GUID), "iid"),
            (["in"], DWORD, "dwClsCtx"),
            (["in"], POINTER(DWORD), "pActivationParams"),
            (["out"], POINTER(POINTER(IUnknown)), "ppInterface"),
        ),
        # HRESULT OpenPropertyStore(
        # [in] DWORD stgmAccess,
        # [out] IPropertyStore **ppProperties);
        COMMETHOD(
            [],
            HRESULT,
            "OpenPropertyStore",
            (["in"], DWORD, "stgmAccess"),
            (["out"], POINTER(POINTER(IPropertyStore)), "ppProperties"),
        ),
        # HRESULT GetId([out] LPWSTR *ppstrId);
        COMMETHOD([], HRESULT, "GetId", (["out"], POINTER(LPWSTR), "ppstrId")),
        # HRESULT GetState([out] DWORD *pdwState);
        COMMETHOD([], HRESULT, "GetState", (["out"], POINTER(DWORD), "pdwState")),
    )


class IMMDeviceCollection(IUnknown):
    _iid_ = GUID("{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}")
    _methods_ = (
        # HRESULT GetCount([out] UINT *pcDevices);
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pcDevices")),
        # HRESULT Item([in] UINT nDevice, [out] IMMDevice **ppDevice);
        COMMETHOD(
            [],
            HRESULT,
            "Item",
            (["in"], UINT, "nDevice"),
            (["out"], POINTER(POINTER(IMMDevice)), "ppDevice"),
        ),
    )


class IMMDeviceEnumerator(IUnknown):
    _iid_ = GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    _methods_ = (
        # HRESULT EnumAudioEndpoints(
        # [in] EDataFlow dataFlow,
        # [in] DWORD dwStateMask,
        # [out] IMMDeviceCollection **ppDevices);
        COMMETHOD(
            [],
            HRESULT,
            "EnumAudioEndpoints",
            (["in"], DWORD, "dataFlow"),
            (["in"], DWORD, "dwStateMask"),
            (["out"], POINTER(POINTER(IMMDeviceCollection)), "ppDevices"),
        ),
        # HRESULT GetDefaultAudioEndpoint(
        # [in] EDataFlow dataFlow,
        # [in] ERole role,
        # [out] IMMDevice **ppDevice);
        COMMETHOD(
            [],
            HRESULT,
            "GetDefaultAudioEndpoint",
            (["in"], DWORD, "dataFlow"),
            (["in"], DWORD, "role"),
            (["out"], POINTER(POINTER(IMMDevice)), "ppDevices"),
        ),
        # HRESULT GetDevice(
        # [in] LPCWSTR pwstrId,
        # [out] IMMDevice **ppDevice);
        COMMETHOD(
            [],
            HRESULT,
            "GetDevice",
            (["in"], LPCWSTR, "pwstrId"),
            (["out"], POINTER(POINTER(IMMDevice)), "ppDevice"),
        ),
        # HRESULT RegisterEndpointNotificationCallback(
        # [in] IMMNotificationClient *pClient);
        COMMETHOD([], HRESULT, "NotImpl1"),
        # HRESULT UnregisterEndpointNotificationCallback(
        # [in] IMMNotificationClient *pClient);
        COMMETHOD([], HRESULT, "NotImpl2"),
    )
