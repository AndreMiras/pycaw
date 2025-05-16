from ctypes import HRESULT, POINTER
from ctypes.wintypes import BOOL, DWORD, LPCWSTR
from comtypes import COMMETHOD, GUID, IUnknown

from pycaw.api.mmdeviceapi import PROPERTYKEY
from pycaw.api.mmdeviceapi.depend import PROPVARIANT


class IPolicyConfig(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID("{f8679f50-850a-41cf-9c72-430f290290c8}")
    _methods_ = (
        COMMETHOD([], HRESULT, "Unused1"),
        COMMETHOD([], HRESULT, "Unused2"),
        COMMETHOD([], HRESULT, "Unused3"),
        COMMETHOD([], HRESULT, "Unused4"),
        COMMETHOD([], HRESULT, "Unused5"),
        COMMETHOD([], HRESULT, "Unused6"),
        COMMETHOD([], HRESULT, "Unused7"),
        COMMETHOD([], HRESULT, "Unused8"),
        COMMETHOD(
            [],
            HRESULT,
            "GetPropertyValue",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], POINTER(PROPERTYKEY), "pKey"),
            (["out"], POINTER(PROPVARIANT), "pValue"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SetPropertyValue",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], POINTER(PROPERTYKEY), "pKey"),
            (["in"], POINTER(PROPVARIANT), "pValue"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SetDefaultEndpoint",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], DWORD, "role"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SetEndpointVisibility",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], BOOL, "bVisible"),
        ),
    )
