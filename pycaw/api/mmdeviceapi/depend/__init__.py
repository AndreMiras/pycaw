from ctypes import HRESULT, POINTER
from ctypes.wintypes import DWORD, LPWSTR

from comtypes import COMMETHOD, GUID, IUnknown

from .structures import PROPERTYKEY, PROPVARIANT


class IPropertyStore(IUnknown):
    _iid_ = GUID('{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}')
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
