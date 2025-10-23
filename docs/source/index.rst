Pycaw - Python Core Audio Windows
===================================

.. image:: https://github.com/AndreMiras/pycaw/actions/workflows/tests.yml/badge.svg
   :target: https://github.com/AndreMiras/pycaw/actions/workflows/tests.yml
   :alt: Tests

.. image:: https://coveralls.io/repos/github/AndreMiras/pycaw/badge.svg?branch=develop
   :target: https://coveralls.io/github/AndreMiras/pycaw?branch=develop
   :alt: Coverage

.. image:: https://badge.fury.io/py/pycaw.svg
   :target: https://badge.fury.io/py/pycaw
   :alt: PyPI version

Pycaw is a Python library designed exclusively for controlling audio devices on **Windows** systems.
It allows programmatic access to audio sessions, volume control, and sound device management on the Windows platform.

.. warning::
   **Windows Only**: Pycaw does not support macOS or Linux.
   It is built specifically for Windows using Core Audio APIs.
   If you're looking for similar functionality on other platforms, you'll need alternative libraries.

Installation
------------

Latest stable release::

    pip install pycaw

Development branch::

    pip install https://github.com/AndreMiras/pycaw/archive/develop.zip

System requirements::

    choco install visualcpp-build-tools

Quick Start
-----------

.. code-block:: python

    from pycaw.pycaw import AudioUtilities
    device = AudioUtilities.GetSpeakers()
    volume = device.EndpointVolume
    print(f"Audio output: {device.FriendlyName}")
    print(f"- Muted: {bool(volume.GetMute())}")
    print(f"- Volume level: {volume.GetMasterVolumeLevel()} dB")
    print(f"- Volume range: {volume.GetVolumeRange()[0]} dB - {volume.GetVolumeRange()[1]} dB")
    volume.SetMasterVolumeLevel(-20.0, None)

Table of Contents
-----------------

.. toctree::
   :maxdepth: 2
   :caption: User Guide

   quickstart
   examples/index

.. toctree::
   :maxdepth: 2
   :caption: API Reference

   api/index

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
