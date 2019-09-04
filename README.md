# pycaw

[![Build Status](https://travis-ci.org/AndreMiras/pycaw.svg?branch=develop)](https://travis-ci.org/AndreMiras/pycaw)
[![PyPI version](https://badge.fury.io/py/pycaw.svg)](https://badge.fury.io/py/pycaw)

Python Core Audio Windows Library, working for both Python2 and Python3.

## Install

Latest stable release:
```bash
pip install pycaw
```

Development branch:
```bash
pip install https://github.com/AndreMiras/pycaw/archive/develop.zip
```

System requirements:
```bash
choco install visualcppbuildtools
```

## Usage

```Python
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volume.GetMute()
volume.GetMasterVolumeLevel()
volume.GetVolumeRange()
volume.SetMasterVolumeLevel(-20.0, None)
```

See more in the [examples](examples/) directory.

## Tests

See in the [tests](tests/) directory.
