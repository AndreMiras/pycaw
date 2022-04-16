# Change Log

## [Unreleased]
    - Fixed GetAllDevices() COMError, refs #15, #28 and #30 (@micolous and @reversefold)
    - New IAudioSessionEvents callbacks support, refs #27, #36 (@TurboAnonym)
    - IAudioSessionControl GetState fix, refs #32, #37 (@TurboAnonym)
    - Adding AudioUtilities.GetMicrophone(), refs #39 (@alebzk)
    - Reorganize / Split the pycaw.pycaw file in modules and subpackages, refs #38 (@TurboAnonym)
    - OnSessionCreated support + Wrapper for callbacks, refs #40 (@TurboAnonym)
    - "Magic" module, easy session control, refs #42 (@TurboAnonym)

## [20190904]
    - Fixed enum34 dependency, refs #17 (@mmxfguerin)

## [20181226]
    - GetVolumeStepInfo() interface fixes
    - Added AudioController class example, refs #4 (@lorenzsj)
    - IAudioEndpointVolumeCallback::OnNotify support, refs #10, #11 (@csevast)
    - Setup (limited) continuous testing, refs #12

## [20160929]
    - Fixed crash on print AudioDevice & AudioSession on Python3
    - Fixed GetAllSessions() reliability, refs #1

## [20160508]
    - Fixed enum requirement
    - Unit tested examples:
        - audio_endpoint_volume_example
        - simple_audio_volume_example
        - volume_by_process_example
    - Added Tox testing framework support
    - Added pyflakes passive checker to tests
    - Testing style convention using pep8
    - Ported code to Python3

## [20160507]
    - Implemented new interfaces:
        - PROPVARIANT
        - IPropertyStore
        - SimpleAudioVolume
        - IAudioClient
    - Added GetAllDevices() helper method
    - Added code examples:
        - audio_endpoint_volume_example
        - simple_audio_volume_example
        - volume_by_process_example

## [20160424]
    - Initial release
