Quick Start Guide
=================

This guide will help you get started with Pycaw.

Requirements
------------

- **Operating System**: Windows (7, 8, 10, 11, Server editions)
- **Python**: 3.8 or higher
- **Visual C++ Build Tools**: Required for comtypes compilation

Installation
------------

Install from PyPI::

    pip install pycaw

Install development version::

    pip install https://github.com/AndreMiras/pycaw/archive/develop.zip

Basic Usage
-----------

Getting the Default Speakers
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code-block:: python

    from pycaw.pycaw import AudioUtilities

    # Get the default speakers
    device = AudioUtilities.GetSpeakers()
    print(f"Default speakers: {device.FriendlyName}")

Controlling Volume
~~~~~~~~~~~~~~~~~~

.. code-block:: python

    from pycaw.pycaw import AudioUtilities

    device = AudioUtilities.GetSpeakers()
    volume = device.EndpointVolume

    # Get current volume info
    current_volume = volume.GetMasterVolumeLevel()
    print(f"Current volume: {current_volume} dB")

    # Get volume range
    vol_range = volume.GetVolumeRange()
    print(f"Volume range: {vol_range[0]} dB to {vol_range[1]} dB")

    # Set volume to -20 dB
    volume.SetMasterVolumeLevel(-20.0, None)

    # Check mute status
    is_muted = volume.GetMute()
    print(f"Muted: {is_muted}")

    # Mute/unmute
    volume.SetMute(1, None)  # Mute
    volume.SetMute(0, None)  # Unmute

Next Steps
----------

- See :doc:`examples/index` for more usage examples
- Check the :doc:`api/index` for complete API reference
