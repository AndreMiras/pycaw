"""
Get and set access to master volume example.
"""
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def main():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    print "volume.GetMute():", volume.GetMute()
    print "volume.GetMasterVolumeLevel():", volume.GetMasterVolumeLevel()
    print "volume.GetVolumeRange():", volume.GetVolumeRange()
    print "volume.SetMasterVolumeLevel()"
    volume.SetMasterVolumeLevel(-20.0, None)
    print "volume.GetMasterVolumeLevel():", volume.GetMasterVolumeLevel()


if __name__ == "__main__":
    main()
