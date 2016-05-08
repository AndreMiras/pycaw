"""
Verifies examples run as expected.
"""
import sys
import unittest
from StringIO import StringIO
from contextlib import contextmanager
from examples import audio_endpoint_volume_example


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

    def test_audio_endpoint_volume_example(self):
        with captured_output() as (out, err):
            audio_endpoint_volume_example.main()
        output = out.getvalue()
        lines = output.split("\n")
        self.assertEqual(lines[0], 'volume.GetMute(): 0')
        self.assertEqual(lines[1], 'volume.GetMasterVolumeLevel(): -20.0')
        self.assertEqual(lines[2], 'volume.GetVolumeRange(): (-95.25, 0.0, 0.75)')
        self.assertEqual(lines[3], 'volume.SetMasterVolumeLevel()')
        self.assertEqual(lines[4], 'volume.GetMasterVolumeLevel(): -20.0')

if __name__ == '__main__':
    unittest.main()
