from ctypes import HRESULT, POINTER, c_longlong, Structure, c_int, c_bool
from ctypes.wintypes import BOOL, DWORD, LPCWSTR
from comtypes import COMMETHOD, GUID, IUnknown

from pycaw.api.audioclient import WAVEFORMATEX
from pycaw.api.mmdeviceapi import PROPERTYKEY
from pycaw.api.mmdeviceapi.depend import PROPVARIANT

REFERENCE_TIME = c_longlong


class DeviceSharedMode(Structure):
    _fields_ = [
        ("Mode", c_int),
        ("bIsEventDriven", c_bool),
    ]


class IPolicyConfig(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID("{f8679f50-850a-41cf-9c72-430f290290c8}")
    _methods_ = (
        COMMETHOD(
            [], 
            HRESULT, 
            "GetMixFormat",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["out"], POINTER(POINTER(WAVEFORMATEX)), "ppDeviceFormat"),
        ),
        COMMETHOD(
            [], 
            HRESULT, 
            "GetDeviceFormat", 
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], BOOL, "bDefault"), 
            (["out"], POINTER(POINTER(WAVEFORMATEX)), "ppDeviceFormat"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "ResetDeviceFormat",
            (["in"], LPCWSTR, "pwstrDeviceId"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SetDeviceFormat",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], POINTER(WAVEFORMATEX), "pEndpointFormat"),
            (["in"], POINTER(WAVEFORMATEX), "pMixFormat"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetProcessingPeriod",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], BOOL, "bDefault"),
            (["out"], POINTER(REFERENCE_TIME), "hnsDefaultDevicePeriod"),
            (["out"], POINTER(REFERENCE_TIME), "hnsMinimumDevicePeriod"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SetProcessingPeriod",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], POINTER(REFERENCE_TIME), "hnsDevicePeriod"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetShareMode",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["out"], POINTER(DeviceSharedMode), "pMode"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SetShareMode",
            (["in"], LPCWSTR, "pwstrDeviceId"),
            (["in"], POINTER(DeviceSharedMode), "pMode"),
        ),
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
