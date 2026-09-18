"""
Verifies examples run as expected.
"""

from unittest import mock

import pytest

from examples import (
    audio_endpoint_volume_example,
    simple_audio_volume_example,
    volume_by_process_example,
)
from pycaw.pycaw import AudioUtilities
from tests.test_core import captured_output


class TestExamples:
    @pytest.mark.skip(reason="Currently failing in the CI")
    def test_audio_endpoint_volume_example(self):
        with captured_output() as (out, err):
            audio_endpoint_volume_example.main()
        output = out.getvalue()
        lines = output.split("\n")
        assert lines[0].startswith("Device found: ")
        assert lines[1] == "volume.GetMute(): 0"
        assert lines[2] == "volume.GetMasterVolumeLevel(): -20.0"
        assert lines[3] == "volume.GetVolumeRange(): (-95.25, 0.0, 0.75)"
        assert lines[4] == "volume.SetMasterVolumeLevel()"
        assert lines[5] == "volume.GetMasterVolumeLevel(): -20.0"

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
