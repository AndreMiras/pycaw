from unittest import mock

from pycaw.magic import MagicApp, MagicManager, MagicSession


def patch_atexit_register():
    """Prevent MagicManager.clean_up() call as it seems to misbehave with tests."""
    return mock.patch("atexit.register")


class TestMagic:
    def test_init(self):
        app_execs = {"msedge.exe"}
        # prevent MagicManager.clean_up() call as it seems to misbehave with tests
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
