import time

from pycaw.magic import MagicApp

# non opp approach:


def handle_all(*args):
    print("callback")
    print(args)


magic = MagicApp("msedge.exe",
                 volume_callback=handle_all,
                 mute_callback=handle_all,
                 state_callback=handle_all,
                 session_callback=handle_all)

magic.volume = 0.1

print(magic.volume)

magic.volume = 0.9

print(magic.volume)

time.sleep(5)

magic.mute = False

print(magic.mute)

magic.mute = True

print(magic.mute)


# opp approach:

# class MyCustomApp(MagicApp):
#     super.init("msedge.exe",
#                volume_callback=self.custom_volume_callback,
#                mute_callback=self.....,
#                state_callback=self.....,
#                session_callback=self.....)

#     def custom_volume_callback(volume):
#         print(volume)

#         print(self.mute)

#         self.mute = True

#         print(self.mute)
