"""
Adjusting volume of left channel using IChannelAudioVolume.
"""

from pycaw.pycaw import AudioUtilities


def main():
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session.channelAudioVolume()
        print(f"Session {session}")
        count = volume.GetChannelCount()
        volumes = [volume.GetChannelVolume(i) for i in range(count)]
        print(f"    volumes = {volumes}")
        if count == 2:
            volume.SetChannelVolume(0, 0.1, None)
            print("    Set the volume of left channel to 0.5!")


if __name__ == "__main__":
    main()
