"""
Verifies examples run as expected.
"""
import sys
import pytest
import unittest
from contextlib import contextmanager
from pycaw.pycaw import AudioUtilities
from examples import audio_endpoint_volume_example
from examples import simple_audio_volume_example
from examples import volume_by_process_example
try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO


@contextmanager
def captured_output():
    new_out, new_err = StringIO(), StringIO()
    old_out, old_err = sys.stdout, sys.stderr
    try:
        sys.stdout, sys.stderr = new_out, new_err
        yield sys.stdout, sys.stderr
    finally:
        sys.stdout, sys.stderr = old_out, old_err


class TestExamples(unittest.TestCase):

    @pytest.mark.skip(reason="Currently failing in the CI")
    def test_audio_endpoint_volume_example(self):
        with captured_output() as (out, err):
            audio_endpoint_volume_example.main()
        output = out.getvalue()
        lines = output.split("\n")
        self.assertEqual(
            lines[0], 'volume.GetMute(): 0')
        self.assertEqual(
            lines[1], 'volume.GetMasterVolumeLevel(): -20.0')
        self.assertEqual(
            lines[2], 'volume.GetVolumeRange(): (-95.25, 0.0, 0.75)')
        self.assertEqual(
            lines[3], 'volume.SetMasterVolumeLevel()')
        self.assertEqual(
            lines[4], 'volume.GetMasterVolumeLevel(): -20.0')

    def test_simple_audio_volume_example(self):
        with captured_output() as (out, err):
            simple_audio_volume_example.main()
        output = out.getvalue()
        lines = output.strip().split("\n")
        sessions = AudioUtilities.GetAllSessions()
        self.assertEqual(len(lines), len(sessions))
        for line in lines:
            self.assertTrue(
                'volume.GetMute(): 0' in line or
                'volume.GetMute(): 1' in line)

    def test_volume_by_process_example(self):
        volume_by_process_example.main()
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() == "chrome.exe":
                self.assertEqual(volume.GetMute(), 0)
            else:
                self.assertEqual(volume.GetMute(), 1)


if __name__ == '__main__':
    unittest.main()
