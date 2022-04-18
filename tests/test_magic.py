from pycaw.magic import MagicApp


class TestMagic:
    def test_init(self):
        app_execs = {"msedge.exe"}
        magic = MagicApp(app_execs)
        assert magic.app_execs == app_execs
