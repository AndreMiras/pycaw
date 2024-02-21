"""
This file contains example usage of MMNotificationClient, the following
callbacks are implemented:

:: Gets called when the state of an audio endpoint device has changed
IMMNotificationClient.OnDeviceStateChanged()
-> on_device_state_changed()

:: Gets called when the value of an audio endpoint property has changed
IMMNotificationClient.OnPropertyValueChanged()
-> on_property_value_changed()

https://learn.microsoft.com/en-us/windows/win32/api/mmdeviceapi/nn-mmdeviceapi-immnotificationclient

"""

import time

from comtypes import GUID, COMError
from comtypes.automation import VT_BLOB
from comtypes.persist import STGM_READ

from pycaw.callbacks import MMNotificationClient
from pycaw.utils import AudioUtilities

known_keys = {
    # Recording
    "{24DBB0FC-9311-4B3D-9CF0-18FF155639D4} 0": "Playback through this device",
    "{24DBB0FC-9311-4B3D-9CF0-18FF155639D4} 1": "Listen to this device",
    # Playback & Recording
    "{9855C4CD-DF8C-449C-A181-8191B68BD06C} 0": "Volume",
    "{9855C4CD-DF8C-449C-A181-8191B68BD06C} 1": "Device muted",
}


class Client(MMNotificationClient):
    def __init__(self):
        self.enumerator = AudioUtilities.GetDeviceEnumerator()

    def on_device_state_changed(self, device_id, new_state, new_state_id):
        print(f"on_device_state_changed: {device_id} {new_state} {new_state_id}")

    def on_property_value_changed(self, device_id, property_struct, fmtid, pid):
        key = f"{fmtid} {pid}"

        value = self._find_property(device_id, fmtid, pid)
        print(
            f"on_property_value_changed: key={key} "
            f"purpose=\"{known_keys.get(key, '?')}\" value={value}",
        )

    def _find_property(self, device_id: str, fmtid: GUID, pid: int) -> str | None:
        """Helper function to find the value of a property"""
        dev = self.enumerator.GetDevice(device_id)
        store = dev.OpenPropertyStore(STGM_READ)
        if store is None:
            print("no store")
            return

        search_value = bytes(fmtid)
        for j in range(store.GetCount()):
            try:
                pk = store.GetAt(j)

                if not (bytes(pk.fmtid) == search_value and pk.pid == pid):
                    continue

                value = store.GetValue(pk)
                if value.vt == VT_BLOB:
                    return bytes(value).hex(" ")

                return value.GetValue()
            except COMError as exc:
                print(
                    f"COMError attempting to get property {j} "
                    f"from device {dev}: {exc}"
                )
                continue


def add_callback():
    cb = Client()
    enumerator = AudioUtilities.GetDeviceEnumerator()
    enumerator.RegisterEndpointNotificationCallback(cb)
    print("registered")

    try:
        # wait for callbacks
        time.sleep(300)
    except KeyboardInterrupt:
        pass
    finally:
        print("unregistering")
        enumerator.UnregisterEndpointNotificationCallback(cb)


if __name__ == "__main__":
    add_callback()
