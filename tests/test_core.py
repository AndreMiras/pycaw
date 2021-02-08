"""
Verifies core features run as expected.
"""
from __future__ import print_function
import _ctypes
import sys
import unittest
from unittest import mock
from contextlib import contextmanager
from pycaw.pycaw import AudioDeviceState, AudioUtilities
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
        """Makes sure printing a session doesn't crash."""
        with captured_output() as (out, err):
            sessions = AudioUtilities.GetAllSessions()
            print("sessions: %s" % sessions)
            for session in sessions:
                print("session: %s" % session)
                print("session.Process: %s" % session.Process)

    def test_device_unicode(self):
        """Makes sure printing a device doesn't crash."""
        with captured_output() as (out, err):
            devices = AudioUtilities.GetAllDevices()
            print("devices: %s" % devices)
            for device in devices:
                print("device: %s" % device)

    def test_device_failed_properties(self):
        """Test that failed properties do not crash the script"""
        dev = mock.Mock()
        dev.GetId = mock.Mock(return_value="id")
        dev.GetState = mock.Mock(return_value=AudioDeviceState.Active)
        store = mock.Mock()
        store.GetCount = mock.Mock(return_value=1)
        store.GetAt = mock.Mock(return_value="pk")
        store.GetValue = mock.Mock(
            side_effect=_ctypes.COMError(None, None, None)
        )
        dev.OpenPropertyStore = mock.Mock(return_value=store)
        with captured_output() as (out, err):
            AudioUtilities.CreateDevice(dev)
        self.assertTrue(
            "UserWarning: COMError attempting to get property 0 from device"
            in err.getvalue()
        )

    def test_getallsessions_reliability(self):
        """
        Verifies AudioUtilities.GetAllSessions() is reliable
        even calling it multiple times, refs:
        https://github.com/AndreMiras/pycaw/issues/1
        """
        for _ in range(100):
            sessions = AudioUtilities.GetAllSessions()
            self.assertTrue(len(sessions) > 0)

if __name__ == '__main__':
    unittest.main()
