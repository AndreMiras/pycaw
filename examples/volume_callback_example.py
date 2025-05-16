"""
IAudioEndpointVolumeCallback.OnNotify() example.
The OnNotify() callback method gets called on volume change.
"""

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
    device = AudioUtilities.GetSpeakers()
    volume = device.EndpointVolume
    callback = AudioEndpointVolumeCallback()
    volume.RegisterControlChangeNotify(callback)
    for _ in range(3):
        volume.SetMute(0, None)
        volume.SetMute(1, None)


if __name__ == "__main__":
    main()
