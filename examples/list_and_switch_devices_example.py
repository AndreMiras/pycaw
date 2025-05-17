"""
Example to list and switch devices.
"""

from pycaw.constants import EDataFlow, DEVICE_STATE
from pycaw.pycaw import AudioUtilities
from pycaw.utils import AudioDevice
import warnings


def get_active_output_devices():
    with warnings.catch_warnings():  # suppress COMError warnings
        warnings.simplefilter("ignore", UserWarning)
        return AudioUtilities.GetAllDevices(data_flow=EDataFlow.eRender.value,
                                            device_state=DEVICE_STATE.ACTIVE.value)


def get_default_device():
    return AudioUtilities.GetSpeakers()


def set_default_device(device: AudioDevice):
    AudioUtilities.SetDefaultDevice(device.id)


if __name__ == "__main__":
    # List devices
    print("List of available output devices (* = default): ")
    active_output_devices = get_active_output_devices()
    default_device = get_default_device()
    other_device = None
    for device in active_output_devices:
        if device.id == default_device.id:
            print(f" * {device.FriendlyName}")
        else:
            print(f"   {device.FriendlyName}")
            other_device = device

    if other_device is not None:
        # Change default to other device
        print(f"Changing default device to {other_device.FriendlyName}...")
        set_default_device(other_device)

        # List devices again
        print("Updated list of available output devices (* = default): ")
        active_output_devices = get_active_output_devices()
        default_device = get_default_device()
        for device in active_output_devices:
            if device.id == default_device.id:
                print(f"* {device.FriendlyName}")
            else:
                print(f"  {device.FriendlyName}")
