# Pycaw (Python Core Audio Windows)

[![Tests](https://github.com/AndreMiras/pycaw/actions/workflows/tests.yml/badge.svg)](https://github.com/AndreMiras/pycaw/actions/workflows/tests.yml)
[![Coverage Status](https://coveralls.io/repos/github/AndreMiras/pycaw/badge.svg?branch=main)](https://coveralls.io/github/AndreMiras/pycaw?branch=main)
[![PyPI release](https://github.com/AndreMiras/pycaw/workflows/PyPI%20release/badge.svg)](https://github.com/AndreMiras/pycaw/actions/workflows/pypi-release.yml)
[![PyPI version](https://badge.fury.io/py/pycaw.svg)](https://badge.fury.io/py/pycaw)
[![Documentation](https://img.shields.io/badge/docs-github%20pages-blue)](https://andremiras.github.io/pycaw/)


Pycaw is a Python library designed exclusively for controlling audio devices on **Windows** systems.
It allows programmatic access to audio sessions, volume control, and sound device management on the Windows platform.

> Note: Pycaw does not support macOS or Linux.
> It is built specifically for Windows using Core Audio APIs.
> If you're looking for similar functionality on other platforms, you'll need alternative libraries.

Pycaw currently supports Python 3.10 through 3.14 on Windows. Supported
Python versions are stable, non-EOL releases that have passed the project's
Windows test matrix; newer Python releases are added after compatibility
validation.


## Install

Latest stable release:
```bat
pip install pycaw
```

Development branch:
```bat
pip install https://github.com/AndreMiras/pycaw/archive/main.zip
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
# or work with the 0-100 scale the Windows volume mixer uses:
print(f"- Volume: {device.volume_percent:.0f}%")
device.volume_percent = 50
```

See more in the [examples](examples/) directory or visit the [documentation](https://andremiras.github.io/pycaw/).

## Tests

Create the locked contributor environment and run the compatibility suite. uv
uses `.python-version` to install/select Python 3.14 and supplies the locked
tools; Tox handles compatibility isolation.

```bat
uv sync --locked
uv run tox
```

Tox remains responsible for isolating the Python 3.10-3.14 minimum/current
dependency matrix. To run one interpreter's lanes, use for example:

```bat
uv run tox -e py314-minimum,py314-current
```

Documentation and release tools are opt-in groups:

```bat
uv sync --locked --group docs
uv run --group release --no-default-groups python -m build
```

Maintainers with a separately installed Tox and the tox-uv plugin can still run
the equivalent `tox` commands directly. See the [tests](tests/) directory for
the test suite.
