"""
IAudioEndpointVolumeCallback.OnNotify() example.
The OnNotify() callback method gets called on volume change.
"""
from ctypes import POINTER, cast

from comtypes import CLSCTX_ALL, COMObject

from pycaw.pycaw import (
    AudioUtilities,
    IAudioEndpointVolume,
    IAudioEndpointVolumeCallback,
)


class AudioEndpointVolumeCallback(COMObject):
    _com_interfaces_ = [IAudioEndpointVolumeCallback]

    def OnNotify(self, pNotify):
        print("OnNotify callback")


def main():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    callback = AudioEndpointVolumeCallback()
    volume.RegisterControlChangeNotify(callback)
    for i in range(3):
        volume.SetMute(0, None)
        volume.SetMute(1, None)


if __name__ == "__main__":
    main()
