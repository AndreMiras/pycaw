"""
Verifies examples run as expected.
"""

from unittest import mock

from examples import (
    audio_endpoint_volume_example,
    simple_audio_volume_example,
    volume_by_process_example,
)
from pycaw.pycaw import AudioUtilities
from tests.test_core import captured_output


class TestExamples:
    def test_audio_endpoint_volume_example(self):
        volume = mock.Mock(
            spec_set=[
                "GetMute",
                "GetMasterVolumeLevel",
                "GetVolumeRange",
                "SetMasterVolumeLevel",
            ]
        )
        volume.GetMute.return_value = 0
        volume.GetMasterVolumeLevel.side_effect = [-10.0, -20.0]
        volume.GetVolumeRange.return_value = (-96.0, 0.0, 1.5)
        device = mock.Mock(spec_set=["FriendlyName", "EndpointVolume"])
        device.FriendlyName = "Mock speakers"
        device.EndpointVolume = volume

        with mock.patch.object(
            audio_endpoint_volume_example.AudioUtilities,
            "GetSpeakers",
            return_value=device,
        ) as get_speakers:
            with captured_output() as (out, err):
                audio_endpoint_volume_example.main()

        assert out.getvalue().splitlines() == [
            "Device found: Mock speakers",
            "volume.GetMute(): 0",
            "volume.GetMasterVolumeLevel(): -10.0",
            "volume.GetVolumeRange(): (-96.0, 0.0, 1.5)",
            "volume.SetMasterVolumeLevel()",
            "volume.GetMasterVolumeLevel(): -20.0",
        ]
        get_speakers.assert_called_once_with()
        volume.SetMasterVolumeLevel.assert_called_once_with(-20.0, None)

    def test_simple_audio_volume_example(self):
        with captured_output() as (out, err):
            simple_audio_volume_example.main()
        output = out.getvalue()
        lines = output.strip().split("\n")
        sessions = AudioUtilities.GetAllSessions()
        assert len(lines) == len(sessions)
        for line in lines:
            assert "volume.GetMute(): 0" in line or "volume.GetMute(): 1" in line

    def test_volume_by_process_example(self):
        chrome_volume = mock.Mock(spec_set=["SetMute"])
        chrome_process = mock.Mock(spec_set=["name"])
        chrome_process.name.return_value = "chrome.exe"
        chrome = mock.Mock(spec_set=["Process", "SimpleAudioVolume"])
        chrome.Process = chrome_process
        chrome.SimpleAudioVolume = chrome_volume

        other_volume = mock.Mock(spec_set=["SetMute"])
        other_process = mock.Mock(spec_set=["name"])
        other_process.name.return_value = "music.exe"
        other = mock.Mock(spec_set=["Process", "SimpleAudioVolume"])
        other.Process = other_process
        other.SimpleAudioVolume = other_volume

        system_volume = mock.Mock(spec_set=["SetMute"])
        system = mock.Mock(spec_set=["Process", "SimpleAudioVolume"])
        system.Process = None
        system.SimpleAudioVolume = system_volume

        with mock.patch.object(
            volume_by_process_example.AudioUtilities,
            "GetAllSessions",
            return_value=[chrome, other, system],
        ) as get_all_sessions:
            volume_by_process_example.main()

        get_all_sessions.assert_called_once_with()
        chrome_volume.SetMute.assert_called_once_with(0, None)
        other_volume.SetMute.assert_called_once_with(1, None)
        system_volume.SetMute.assert_called_once_with(1, None)
