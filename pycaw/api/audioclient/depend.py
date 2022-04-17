from ctypes import Structure
from ctypes.wintypes import WORD


class WAVEFORMATEX(Structure):
    _fields_ = [
        ("wFormatTag", WORD),
        ("nChannels", WORD),
        ("nSamplesPerSec", WORD),
        ("nAvgBytesPerSec", WORD),
        ("nBlockAlign", WORD),
        ("wBitsPerSample", WORD),
        ("cbSize", WORD),
    ]
