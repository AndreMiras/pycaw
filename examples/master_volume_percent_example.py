"""
Get and set the master volume by percentage example.

`volume_percent` uses the same 0-100 scale as the Windows volume mixer,
unlike `SetMasterVolumeLevel()` which expects decibels.
"""

from pycaw.pycaw import AudioUtilities


def main():
    device = AudioUtilities.GetSpeakers()
    print("Device found: %s" % device.FriendlyName)
    original = device.volume_percent
    print("volume_percent: %.0f%%" % original)
    print("setting volume to 50%")
    device.volume_percent = 50
    print("volume_percent: %.0f%%" % device.volume_percent)
    print("restoring original volume")
    device.volume_percent = original
    print("volume_percent: %.0f%%" % device.volume_percent)


if __name__ == "__main__":
    main()
