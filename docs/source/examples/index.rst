Examples
========

This page provides examples of common Pycaw usage patterns.

All example source files are available in the `examples/ directory <https://github.com/AndreMiras/pycaw/tree/develop/examples>`_ of the repository.

Audio Endpoint Volume Control
------------------------------

Basic master volume control using the default speakers.

.. literalinclude:: ../../../examples/audio_endpoint_volume_example.py
   :language: python
   :linenos:

Volume Control by Process
--------------------------

Control volume for individual applications/processes.

.. literalinclude:: ../../../examples/volume_by_process_example.py
   :language: python
   :linenos:

List and Switch Devices
------------------------

Enumerate available audio devices and switch between them.

.. literalinclude:: ../../../examples/list_and_switch_devices_example.py
   :language: python
   :linenos:

More Examples
-------------

Additional examples are available in the repository:

- **Session Callbacks**: ``session_callback_example.py`` - Event handling for audio session changes
- **Volume Callbacks**: ``volume_callback_example.py`` - Event handling for volume changes
- **Notification Client**: ``notification_client_example.py`` - Device change notifications
- **Channel Audio Volume**: ``channel_audio_volume_example.py`` - Per-channel volume control
- **Simple Audio Volume**: ``simple_audio_volume_example.py`` - Simple session volume control
- **Magic App Example**: ``magic_app_example.py`` - MTA COM initialization for notifications
- **Audio Controller Class**: ``audio_controller_class_example.py`` - Object-oriented wrapper example

You can view all examples on `GitHub <https://github.com/AndreMiras/pycaw/tree/develop/examples>`_.
