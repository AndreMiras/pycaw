"""
Note
----
'import pycaw.magic' must be generally at the topmost.
To be more specific:
It needs to be imported before any other pycaw or comtypes import.


Reserved Atrributes
-------------------
Note that certain methods and attributes are reserved for the magic module.
    Please look into the source code of MagicApp for more information.
But to avoid conflicts now and in the future, i recommend using
a prefix for each of your custom methods and attributes.


Features
--------
Instantiate a new MagicApp with one or more app executables:

magic = MagicApp({"msedge.exe", "another.exe"})

--------

you could also inherit from MagicApp and create customized callbacks:

class MyCustomApp(MagicApp):
    def __init__(self, app_execs):
        super().__init__(app_execs,
                         volume_callback=self.custom_volume_callback,
                         mute_callback=self...,
                         state_callback=self...,
                         session_callback=self...)

    def custom_volume_callback(self, volume):
        print(volume)
        print(self.mute)
        self.mute = True
        print(self.mute)

mega_magic = MyCustomApp({"msedge.exe"})
"""

import time
from contextlib import suppress

from pycaw.magic import MagicApp


def handle_all(*args):
    print("callback")
    print(args)


magic = MagicApp({"msedge.exe"},
                 volume_callback=handle_all,
                 mute_callback=handle_all,
                 state_callback=handle_all,
                 session_callback=handle_all)


def main():
    with suppress(KeyboardInterrupt):
        for _ in range(5):
            """
            open and close your MagicApp app_exec (msedge.exe)
            and see how it will change the volume as long as
            the app is opened. When you close app_exec it wont change
            the volume and None is printed.

            if you change for example the volume in the Windows sound mixer
            handle_all() is fired.
            """

            if magic.state is None:
                print(f"No session active for: {magic}")
                time.sleep(2)
                continue

            print("Volume:")
            magic.volume = 0.1
            print(magic.volume)
            time.sleep(1)
            magic.volume = 0.9
            print(magic.volume)
            time.sleep(1)

            print(f"{str(magic.state)} {magic.app_execs}")

            print("Mute:")
            magic.mute = True
            print(magic.mute)
            time.sleep(1)
            magic.mute = False
            print(magic.mute)

    print("\nTschüss")


if __name__ == '__main__':
    main()
