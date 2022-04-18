from pycaw.magic import MagicApp, MagicManager, MagicSession


class TestMagic:
    def test_init(self):
        app_execs = {"msedge.exe"}
        magic = MagicApp(app_execs)
        assert magic.app_execs == app_execs


class TestMagicManager:
    def test_magic_session(self):
        assert MagicManager.MagicSessionConfigured is None
        MagicManager.magic_session(MagicSession)
        assert MagicManager.MagicSessionConfigured == (MagicSession, (), {})
