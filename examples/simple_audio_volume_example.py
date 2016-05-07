"""
Per session GetMute() SetMute() using ISimpleAudioVolume.
"""
from pycaw import AudioUtilities, ISimpleAudioVolume


def main():
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        print u"volume.GetMute():", volume.GetMute()
        volume.SetMute(1, None)


if __name__ == "__main__":
    main()
