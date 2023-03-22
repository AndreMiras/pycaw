"""
Verifies examples run as expected.
"""
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
        assert lines[0] == "volume.GetMute(): 0"
        assert lines[1] == "volume.GetMasterVolumeLevel(): -20.0"
        assert lines[2] == "volume.GetVolumeRange(): (-95.25, 0.0, 0.75)"
        assert lines[3] == "volume.SetMasterVolumeLevel()"
        assert lines[4] == "volume.GetMasterVolumeLevel(): -20.0"

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
        volume_by_process_example.main()
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() == "chrome.exe":
                assert volume.GetMute() == 0
            else:
                assert volume.GetMute() == 1
