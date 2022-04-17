"""
Get and set access to master volume example.
"""
from ctypes import POINTER, cast

from comtypes import CLSCTX_ALL

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def main():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    print("volume.GetMute(): %s" % volume.GetMute())
    print("volume.GetMasterVolumeLevel(): %s" % volume.GetMasterVolumeLevel())
    print("volume.GetVolumeRange(): (%s, %s, %s)" % volume.GetVolumeRange())
    print("volume.SetMasterVolumeLevel()")
    volume.SetMasterVolumeLevel(-20.0, None)
    print("volume.GetMasterVolumeLevel(): %s" % volume.GetMasterVolumeLevel())


if __name__ == "__main__":
    main()
