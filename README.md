# Pycaw (Python Core Audio Windows)

[![Tests](https://github.com/AndreMiras/pycaw/actions/workflows/tests.yml/badge.svg)](https://github.com/AndreMiras/pycaw/actions/workflows/tests.yml)
[![Coverage Status](https://coveralls.io/repos/github/AndreMiras/pycaw/badge.svg?branch=develop)](https://coveralls.io/github/AndreMiras/pycaw?branch=develop)
[![PyPI release](https://github.com/AndreMiras/pycaw/workflows/PyPI%20release/badge.svg)](https://github.com/AndreMiras/pycaw/actions/workflows/pypi-release.yml)
[![PyPI version](https://badge.fury.io/py/pycaw.svg)](https://badge.fury.io/py/pycaw)
[![Documentation](https://img.shields.io/badge/docs-github%20pages-blue)](https://andremiras.github.io/pycaw/)


Pycaw is a Python library designed exclusively for controlling audio devices on **Windows** systems.
It allows programmatic access to audio sessions, volume control, and sound device management on the Windows platform.

> Note: Pycaw does not support macOS or Linux.
> It is built specifically for Windows using Core Audio APIs.
> If you're looking for similar functionality on other platforms, you'll need alternative libraries.


## Install

Latest stable release:
```bat
pip install pycaw
```

Development branch:
```bat
pip install https://github.com/AndreMiras/pycaw/archive/develop.zip
```

System requirements:
```bat
choco install visualcpp-build-tools
```

## Usage

```Python
from pycaw.pycaw import AudioUtilities
device = AudioUtilities.GetSpeakers()
volume = device.EndpointVolume
print(f"Audio output: {device.FriendlyName}")
print(f"- Muted: {bool(volume.GetMute())}")
print(f"- Volume level: {volume.GetMasterVolumeLevel()} dB")
print(f"- Volume range: {volume.GetVolumeRange()[0]} dB - {volume.GetVolumeRange()[1]} dB")
volume.SetMasterVolumeLevel(-20.0, None)
```

See more in the [examples](examples/) directory or visit the [documentation](https://andremiras.github.io/pycaw/).

## Tests

See in the [tests](tests/) directory.
