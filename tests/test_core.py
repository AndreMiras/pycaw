"""
Verifies core features run as expected.
"""

import _ctypes
import sys
import warnings
from contextlib import contextmanager
from io import StringIO
from unittest import mock

import pytest

from pycaw.pycaw import AudioDeviceState, AudioUtilities
from pycaw.utils import AudioDevice, AudioSession


@contextmanager
def captured_output():
    new_out, new_err = StringIO(), StringIO()
    old_out, old_err = sys.stdout, sys.stderr
    try:
        sys.stdout, sys.stderr = new_out, new_err
        yield sys.stdout, sys.stderr
    finally:
        sys.stdout, sys.stderr = old_out, old_err


class TestCore:
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
        store.GetValue = mock.Mock(side_effect=_ctypes.COMError(None, None, None))
        dev.OpenPropertyStore = mock.Mock(return_value=store)
        with warnings.catch_warnings(record=True) as w:
            AudioUtilities.CreateDevice(dev)
        assert len(w) == 1
        assert "COMError attempting to get property 0 from device" in str(w[0].message)

    def test_getallsessions_reliability(self):
        """
        Verifies AudioUtilities.GetAllSessions() is reliable
        even calling it multiple times, refs:
        https://github.com/AndreMiras/pycaw/issues/1
        """
        for _ in range(100):
            sessions = AudioUtilities.GetAllSessions()
            assert len(sessions) > 0

    def test_volume_percent(self):
        """
        volume_percent maps to the endpoint volume scalar (0-100 <-> 0.0-1.0)
        and the setter clamps out of range values, refs:
        https://github.com/AndreMiras/pycaw/issues/13
        """
        device = AudioDevice("id", AudioDeviceState.Active, {}, mock.Mock())
        endpoint = mock.Mock()
        endpoint.GetMasterVolumeLevelScalar = mock.Mock(return_value=0.25)
        device._volume = endpoint
        assert device.volume_percent == 25.0
        device.volume_percent = 50
        endpoint.SetMasterVolumeLevelScalar.assert_called_with(0.5, None)
        device.volume_percent = 150
        endpoint.SetMasterVolumeLevelScalar.assert_called_with(1.0, None)
        device.volume_percent = -10
        endpoint.SetMasterVolumeLevelScalar.assert_called_with(0.0, None)

    def test_audio_session_notification_lifecycle(self):
        control = mock.Mock()
        session = AudioSession(control)
        first_callback = mock.sentinel.first_callback
        second_callback = mock.sentinel.second_callback

        assert session._callback is None

        session.register_notification(first_callback)
        control.RegisterAudioSessionNotification.assert_called_once_with(first_callback)
        assert session._callback is first_callback

        session.unregister_notification()
        control.UnregisterAudioSessionNotification.assert_called_once_with(
            first_callback
        )
        assert session._callback is None

        session.unregister_notification()
        control.UnregisterAudioSessionNotification.assert_called_once_with(
            first_callback
        )

        session.register_notification(second_callback)
        assert control.RegisterAudioSessionNotification.call_args_list == [
            mock.call(first_callback),
            mock.call(second_callback),
        ]
        assert session._callback is second_callback

    def test_audio_session_keeps_callback_when_unregister_fails(self):
        control = mock.Mock()
        session = AudioSession(control)
        callback = mock.sentinel.callback
        session.register_notification(callback)
        control.UnregisterAudioSessionNotification.side_effect = RuntimeError(
            "unregister failed"
        )

        with pytest.raises(RuntimeError, match="unregister failed"):
            session.unregister_notification()

        control.UnregisterAudioSessionNotification.assert_called_once_with(callback)
        assert session._callback is callback
