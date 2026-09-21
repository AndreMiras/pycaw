import warnings
from unittest import mock

import psutil
import pytest

from pycaw.magic import MagicApp, MagicManager, MagicSession, _MagicRootSession


def patch_atexit_register():
    """Prevent MagicManager.clean_up() call as it seems to misbehave with tests."""
    return mock.patch("atexit.register")


def make_root_session_receiver(activated=True, volume=0.25, mute=0):
    receiver = mock.Mock(
        spec_set=[
            "_activated",
            "volume",
            "mute",
            "app_exec",
            "magic_app",
            "magic_session",
            "_send_callback",
        ]
    )
    receiver._activated = activated
    receiver.volume = volume
    receiver.mute = mute
    receiver.app_exec = "test.exe"
    receiver.magic_app = mock.sentinel.magic_app
    receiver.magic_session = mock.sentinel.magic_session
    receiver._send_callback = mock.Mock()
    return receiver


class TestMagicRootSession:
    def test_get_app_exec(self):
        receiver = mock.Mock(_ctl2=mock.Mock())
        receiver._ctl2.GetProcessId.return_value = 42

        with mock.patch("pycaw.magic.psutil.Process") as process:
            process.return_value.name.return_value = "test.exe"
            app_exec = _MagicRootSession._get_app_exec(receiver)

        assert app_exec == "test.exe"
        process.assert_called_once_with(42)

    def test_get_app_exec_process_exited(self):
        receiver = mock.Mock(_ctl2=mock.Mock())
        receiver._ctl2.GetProcessId.return_value = 42

        with mock.patch(
            "pycaw.magic.psutil.Process", side_effect=psutil.NoSuchProcess(42)
        ):
            app_exec = _MagicRootSession._get_app_exec(receiver)

        assert app_exec is None

    def test_get_app_exec_system_sounds(self):
        receiver = mock.Mock(_ctl2=mock.Mock())
        receiver._ctl2.GetProcessId.return_value = 0
        receiver._ctl2.IsSystemSoundsSession.return_value = 0

        app_exec = _MagicRootSession._get_app_exec(receiver)

        assert app_exec == "SndVol.exe"

    def test_get_app_exec_unidentified_processless_session(self):
        receiver = mock.Mock(_ctl2=mock.Mock())
        receiver._ctl2.GetProcessId.return_value = 0
        receiver._ctl2.IsSystemSoundsSession.return_value = 1

        with pytest.raises(ValueError, match="unidentified app"):
            _MagicRootSession._get_app_exec(receiver)

    def test_volume_change(self):
        receiver = make_root_session_receiver()
        event_context = mock.sentinel.event_context

        _MagicRootSession.OnSimpleVolumeChanged(receiver, 0.75, 0, event_context)

        assert receiver.volume == 0.75
        assert receiver.mute == 0
        assert receiver._send_callback.call_args_list == [
            mock.call(receiver.magic_app, "volume_callback", event_context, 0.75),
            mock.call(receiver.magic_session, "volume_callback", event_context, 0.75),
        ]

    def test_mute_change(self):
        receiver = make_root_session_receiver()
        event_context = mock.sentinel.event_context

        _MagicRootSession.OnSimpleVolumeChanged(receiver, 0.25, 1, event_context)

        assert receiver.volume == 0.25
        assert receiver.mute == 1
        assert receiver._send_callback.call_args_list == [
            mock.call(receiver.magic_app, "mute_callback", event_context, 1),
            mock.call(receiver.magic_session, "mute_callback", event_context, 1),
        ]

    def test_combined_volume_and_mute_change(self):
        receiver = make_root_session_receiver()
        event_context = mock.sentinel.event_context

        _MagicRootSession.OnSimpleVolumeChanged(receiver, 0.75, 1, event_context)

        assert receiver.volume == 0.75
        assert receiver.mute == 1
        assert receiver._send_callback.call_args_list == [
            mock.call(receiver.magic_app, "volume_callback", event_context, 0.75),
            mock.call(receiver.magic_session, "volume_callback", event_context, 0.75),
            mock.call(receiver.magic_app, "mute_callback", event_context, 1),
            mock.call(receiver.magic_session, "mute_callback", event_context, 1),
        ]

    def test_unchanged_volume_and_mute(self):
        receiver = make_root_session_receiver()

        _MagicRootSession.OnSimpleVolumeChanged(
            receiver, 0.25, 0, mock.sentinel.event_context
        )

        receiver._send_callback.assert_not_called()

    def test_inactive_session_ignores_changes(self):
        receiver = make_root_session_receiver(activated=False)

        _MagicRootSession.OnSimpleVolumeChanged(
            receiver, 0.75, 1, mock.sentinel.event_context
        )

        assert receiver.volume == 0.25
        assert receiver.mute == 0
        receiver._send_callback.assert_not_called()


class TestMagic:
    def test_init(self):
        app_execs = {"msedge.exe"}
        with patch_atexit_register() as m_register:
            magic = MagicApp(app_execs)
        assert magic.app_execs == app_execs
        assert m_register.called


class TestMagicManager:
    def test_magic_session(self):
        assert MagicManager.MagicSessionConfigured is None
        MagicManager.magic_session(MagicSession)
        assert MagicManager.MagicSessionConfigured == (MagicSession, (), {})

    def test_unregister_all(self):
        assert MagicManager.magic_apps is not None
        assert MagicManager.str() == (
            "<MagicManager magic_apps='1' magic_sessions='1'"
            " active_mrs='1' trash_mrs='0'/>"
        )
        MagicManager.unregister_all()
        assert MagicManager.str() == "unactive MagicManager"
        assert hasattr(MagicManager, "magic_apps") is False
        assert hasattr(MagicManager, "magic_sessions") is False
        assert MagicManager.magic_activated is None

    def test_clean_up(self):
        app_execs = {"msedge.exe"}
        with patch_atexit_register(), warnings.catch_warnings(record=True):
            MagicApp(app_execs)
        assert MagicManager.magic_apps is not None
        MagicManager.clean_up()
        assert MagicManager.str() == "unactive MagicManager"
        assert hasattr(MagicManager, "magic_apps") is False
        assert hasattr(MagicManager, "magic_sessions") is False

    def test_activate_magic(self):
        MagicManager.magic_activated = None
        app_execs = {"msedge.exe"}
        with patch_atexit_register() as m_register, warnings.catch_warnings(
            record=True
        ) as w:
            MagicApp(app_execs)
        assert m_register.called
        assert len(w) == 1
        MagicManager.unregister_all()
        with warnings.catch_warnings(record=True) as w:
            MagicManager.activate_magic()
        assert len(w) == 1
        assert "<MagicManager/> was already activated an closed." in str(w[-1].message)
        assert MagicManager.magic_activated is True
        with warnings.catch_warnings(record=True) as w:
            MagicManager.activate_magic()
        assert len(w) == 1
        assert "cannot activate MagicManager. MagicManager is already active!" in str(
            w[-1].message
        )
        assert MagicManager.magic_activated is True
