"""
Verifies core features run as expected.
"""
from __future__ import print_function
import sys
import unittest
from contextlib import contextmanager
from pycaw.pycaw import AudioUtilities
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


class TestCore(unittest.TestCase):

    def test_session_unicode(self):
        """
        Makes sure printing a session doesn't crash.
        """
        with captured_output() as (out, err):
            sessions = AudioUtilities.GetAllSessions()
            print("sessions: %s" % sessions)
            for session in sessions:
                print("session: %s" % session)
                print("session.Process: %s" % session.Process)

    def test_device_unicode(self):
        """
        Makes sure printing a device doesn't crash.
        """
        with captured_output() as (out, err):
            devices = AudioUtilities.GetAllDevices()
            print("devices: %s" % devices)
            for device in devices:
                print("device: %s" % device)


if __name__ == '__main__':
    unittest.main()
