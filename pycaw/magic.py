"""
Note
----
in order for AudioSessionNotification::OnSessionCreated to work
the entry point needs to CoInitialize
the main thread in Multi-threaded Apartment (MTA).

All the other Notification will work in the "default" comtypes mode:
    Single-threaded Apartment (STA)

Its generally true that when you work with COM components,
that to need to match the apartment to the component.
But since we are only accessing COM components,
it works using a MTA client to access both STA and MTA COM objects.

isort:skip_file
"""

# flake8:willaddbacklater noqa
# ____ COM WITH MULTITHREADED APARTMENT ____
import sys

sys.coinit_flags = 0  # noqa: E402

import atexit
import logging
import os
import sys
import warnings
from ctypes import pointer

import psutil
from _ctypes import COMError
from comtypes import GUID, COMObject

from pycaw.api.audioclient import ISimpleAudioVolume
from pycaw.api.audiopolicy import (IAudioSessionControl2, IAudioSessionEvents,
                                   IAudioSessionNotification)
from pycaw.utils import AudioUtilities

log = logging.getLogger('Pycaw Magic')


# class log:
#     @staticmethod
#     def info(msg):
#         print(msg)


class MagicManager(COMObject):
    """
    The 'MagicManager' handles the magic.

    Features
    --------
    -   handles adding and removing sessions from a dictionary.
    -   the main dict 'magic_root_sessions' contains all active sessions.
    -   the sessions are idendified by the GetSessionInstanceIdentifier()
            the iid (InstanceIdentifier)
            is unique for every new session of any app.

    -   If one or more MagicApps are hooked into the MagicManager,
        by creating a new instance of a MagicApp,
        the MagicManager will hand over the requested sessions.
        (Those sessions which match the MagicApp(app_exe) like firefox.exe)
        Multiple apps can be controlled by passing multiple a list to MagicApp:
            MagicApp(["firefox.exe", "vlc.exe"])
            When an app name has multiple audio sessions,
            they will be automatically 'merged'
        Note that when using MagicApp in multi session mode, that getting the
        volume will result in getting the volume of the loudest session.
        Mute and State will return if availble the '1' state.

    -   handles checking if session already exists
        by using the session.InstanceIdentifier
        (sometimes a session fires multiple times OnSessionCreated
        even its already there) TODO: is that actually true?

    -   unregister all (still active) sessions from callback at shutdown

    -   OnSessionCreated is fired by Windows everytime a new session registers
    """

    _com_interfaces_ = (IAudioSessionNotification,)
    magic_activated = False

    @classmethod
    def activate_magic(cls):
        """
        Gets only called once or never. Not on import.
        Depending if MagicApp or MagicSession are used.
        see:
            self.add_magic_app()
            self.magic_session()
        """
        if cls.magic_activated:
            log.warning("cannot activate MagicManager. "
                        "MagicManager is already active!")
            return

        log.info("activate magic")

        # dict with idd -> magic_root_session -> MagicRootSession
        cls.magic_root_sessions = {}

        # set of MagicApp instances
        cls.magic_apps = set()

        # if registered via MagicManager.magic_session()
        # will hold the MagicSession and additional args + kwargs
        cls.MagicSessionConfigured = None
        # dict like
        cls.magic_sessions = {}

        try:
            mgr = AudioUtilities.GetAudioSessionManager()
            log.info("got manager")
        except COMError:
            exit("No speaker connected")

        magic_manager = cls()
        mgr.RegisterSessionNotification(magic_manager)
        # has to get called -
        # to make IAudioSessionNotification::OnSessionCreated working
        sessionEnumerator = mgr.GetSessionEnumerator()

        log.info("MagicManager: registered and activated session notification")

        # Scan for running session and add them to session_manager

        # get all active sessions
        count = sessionEnumerator.GetCount()
        log.info(f"{count} sessions already active")

        # add sessions to session_manager
        for i in range(count):
            log.info("adding session manually")
            ctl = sessionEnumerator.GetSession(i)
            cls.OnSessionCreated(ctl)

        atexit.register(cls.clean_up, mgr, magic_manager)

        cls.magic_activated = True

    @classmethod
    def clean_up(cls, mgr, magic_manager):
        log.info("exiting")
        mgr.UnregisterSessionNotification(magic_manager)
        cls.unregister_all()

    @classmethod
    def OnSessionCreated(cls, ctl):
        ctl2 = ctl.QueryInterface(IAudioSessionControl2)
        iid = ctl2.GetSessionInstanceIdentifier()
        log.info(":: new session")
        if iid not in cls.magic_root_sessions:
            log.info("create a new magic_root_session for:")
            magic_root_session = MagicRootSession(ctl2,
                                                  iid=iid,
                                                  magic_manager=cls)
            log.info(magic_root_session.app_exe)
            cls.magic_root_sessions[iid] = magic_root_session
            # add exe to matching magic app

            if cls.magic_apps:
                new_app_exe = magic_root_session.app_exe
                log.info(f"searching matching magic_app for: {new_app_exe}")
                for magic_app in cls.magic_apps:
                    for app_exe in magic_app.app_exe:
                        if app_exe == new_app_exe:
                            log.info(f"{new_app_exe} matched {magic_app}."
                                     "Adding magic_root_session"
                                     "to magic_app...")
                            magic_app.add_magic_root_session(
                                iid, magic_root_session)
            if cls.MagicSessionConfigured:
                MagicSessionClass, args, kwargs = cls.MagicSessionConfigured
                magic_session = MagicSessionClass(magic_root_session,
                                                  *args, **kwargs)
                cls.magic_sessions[iid] = magic_session

        else:
            warnings.warn(f"{session.Process}\nis already present")

    @classmethod
    def magic_session(cls, MagicSessionClass, *args, **kwargs):
        """
        gets called explicit by the user with a custom MagicSession.
        this will tell the MagicManager to create for all current
        and new sessions a custom MagicSession.
        """
        if not cls.magic_activated:
            cls.activate_magic()
        for iid, magic_root_session in cls.magic_root_sessions.items():
            magic_session = MagicSessionClass(magic_root_session,
                                              *args, **kwargs)
            cls.magic_sessions[iid] = magic_session
        cls.MagicSessionConfigured = (MagicSessionClass, args, kwargs)

    @classmethod
    def add_magic_app(cls, magic_app, apps_exe):
        """
        gets called when a new magic_app is created.
        this will tell the MagicManager add all current and
        new matching sessions by app_exe to the magic_app.
        """
        if not cls.magic_activated:
            cls.activate_magic()
        log.info("trying to find match sessions,"
                 f"which are already active, for: {magic_app}")
        for app_exe in apps_exe:
            for iid, magic_root_session in cls.magic_root_sessions.items():
                if magic_root_session.app_exe == app_exe:
                    log.info(f"{app_exe} matched {magic_app}."
                             "Adding magic_root_session to magic_app...")
                    magic_app.add_magic_root_session(iid, magic_root_session)

        # keep reference to magic_app to check later
        # if new session should be added to this magic_app
        cls.magic_apps.add(magic_app)
        log.info(f"added {magic_app} to watchlist ({len(cls.magic_apps)})")

    @classmethod
    def remove_session(cls, iid, magic_app=None):
        """magic_root_session will get removed because it is expired"""
        magic_root_session = cls.magic_root_sessions.pop(iid, None)

        # unregister callback
        # (magic_root_session -> MagicRootSession -> IAudioSessionEvents)
        if magic_root_session:
            magic_root_session.unregister_notification()

        # delete iid also from magic_app or magic_session
        # TODO: add remove handle callback or keep only way with:
        # if new_state_id == 3: ?
        cls.magic_sessions.pop(iid, None)
        if magic_app:
            magic_app.magic_root_sessions.pop(iid, None)

    @classmethod
    def unregister_all(cls):
        # since cls.active can always change (multithreading)
        # instead of a for loop i present you the while loop:
        # while cls.active contains any items,
        # they get popped and unregistered
        while cls.magic_root_sessions:
            # ________ REMOVES 1 ITEM FROM LIST ________
            _, magic_root_session = cls.magic_root_sessions.popitem()

            try:
                # could raise an error if a session is closed
                # between popitem() and unregister_notification()
                magic_root_session.unregister_notification()
                log.info(":: :: :: unregister Session")
            except COMError as e:
                # caused in pycaw.utils.AudioSession
                # self._ctl.UnregisterAudioSessionNotification(self._callback)
                # 'element not found'
                warnings.warn(
                    f"\nSession and Mixer closed simultaneous?\n{str(e)}")


def for_sessions(func):
    """Decorator for looping through sessions in MagicApp."""
    def wrapper(self, *args):
        if self.magic_root_sessions:
            rv = [
                func(self, magic_root_session, *args)
                for magic_root_session in self.magic_root_sessions.values()
                ]
            # return if multiple sessions the one
            # with mute true or the highest volume
            rv = max(rv)
            return rv
        else:
            return None
    return wrapper


class MagicApp():
    """
    When instantiated with atleast one app_exe name,
    will be able to get/ set volume etc, also if the
    session is created after initialize.
    """
    guid = pointer(GUID('{E0BD1A40-9624-44FC-A607-2ED4F00B1CC4}'))

    def __init__(self,
                 app_exe,
                 volume_callback=None,
                 mute_callback=None,
                 state_callback=None,
                 session_callback=None):

        # normalize app_exe
        if type(app_exe) == str:
            app_exe = (app_exe,)
        self.app_exe = set(app_exe)

        # latest dict of matching sessions
        self.magic_root_sessions = {}

        # callbacks
        self.volume_callback = volume_callback
        self.mute_callback = mute_callback
        self.state_callback = state_callback
        self.session_callback = session_callback

        log.info(str(self))
        MagicManager.add_magic_app(self, app_exe)

    def add_magic_root_session(self, iid, magic_root_session):
        """called by MagicManager, when a new matching session is found"""
        self.magic_root_sessions[iid] = magic_root_session
        # tells the magic_root_session to create a connection
        # for callbacks.
        magic_root_session.use_magic_app(self)
        if self.session_callback:
            self.session_callback(magic_root_session)

    def __str__(self):
        return (f"Magic: {self.__class__} Registered for: {self.app_exe}"
                f"Controls {len(self.magic_root_sessions)} sessions")

    # easy control:
    @property
    @for_sessions
    def state(self, magic_root_session):
        return magic_root_session.state

    @property
    @for_sessions
    def volume(self, magic_root_session):
        return magic_root_session.volume

    @volume.setter
    @for_sessions
    def volume(self, magic_root_session, volume):
        magic_root_session._sav.SetMasterVolume(volume, self.guid)

    @property
    @for_sessions
    def mute(self, magic_root_session):
        return magic_root_session.mute

    @mute.setter
    @for_sessions
    def mute(self, magic_root_session, mute):
        magic_root_session._sav.SetMute(mute, self.guid)

    def toogle_mute(self):
        new_mute = not self.mute
        self.mute = new_mute
        return new_mute


class MagicSession():
    """
    When activated and passed via
    MagicManager.magic_session(inherited_class_of_MagicSession)
    will be created for each new and current session.
    Dict of all iid -> MagicSession:
        MagicManager.magic_sessions
    """
    guid = pointer(GUID('{34482A3D-37DD-40E3-BB37-63A16036C87C}'))

    def __init__(self,
                 magic_root_session,
                 volume_callback=None,
                 mute_callback=None,
                 state_callback=None):

        self.magic_root_session = magic_root_session
        # tells the magic_root_session to create a connection
        # for callbacks:
        self.magic_root_session.use_magic_session(self)

        # callbacks
        self.volume_callback = volume_callback
        self.mute_callback = mute_callback
        self.state_callback = state_callback

        log.info(str(self))

    def __str__(self):
        return (f"MagicSession: {self.__class__}"
                f"for app: {self.magic_root_session.app_exe}")

    # easy control:
    @property
    def state(self):
        return self.magic_root_session.state

    @property
    def volume(self):
        return self.magic_root_session.volume

    @volume.setter
    def volume(self, volume):
        self.magic_root_session._sav.SetMasterVolume(volume, self.guid)

    @property
    def mute(self):
        return self.magic_root_session.mute

    @mute.setter
    def mute(self, mute):
        self.magic_root_session._sav.SetMute(mute, self.guid)

    def toogle_mute(self):
        new_mute = not self.mute
        self.mute = new_mute
        return new_mute


class MagicRootSession(COMObject):
    """Base session control with callback functionality"""

    _com_interfaces_ = (IAudioSessionEvents,)

    # ======= DECODE RETURNED INT VALUE =======
    # see audiosessiontypes.h
    AudioSessionState = (
        "Inactive",
        "Active",
        "Expired"
    )

    def __init__(self, audio_session_control2, iid=None, magic_manager=None):
        self._ctl2 = audio_session_control2
        self.app_exe = self._get_app_exe()
        self._sav = None
        self.magic_manager = magic_manager
        self.register_notification()
        if not iid:
            iid = self._ctl2.GetSessionInstanceIdentifier()
        self.iid = iid

        self.magic_app = None
        self.magic_session = None

        self.volume = None
        self.mute = None
        self.state = None
        self.state_id = None

    def __str__(self):
        return f"{self.__class__} for: {self.app_exe}"

    # TODO:
    # should only one magic_app at the time be able to
    # control and view this magic_root_session?
    # if so strict rules need to apply on use_magic_app()
    #
    # -------------------
    # If multiple magic_app would be allowed:
    # Looping through multiple magic_app and checking for
    # callbacks in each would cost lots of time right?
    # any solution?
    #
    # -------------------
    # WARNING:
    # the Ministry of Magic doesnt stop you in the current
    # state to use multiple magic_app for this session.
    # Not implemented!
    # -----
    # To stop multiple apps this use_magic_app must send a
    # false signal up back to magic_app wich then would
    # unregister this magic_app... or it could use this sessions
    # without callbacks.
    def use_magic_app(self, magic_app):
        self.magic_app = magic_app
        self._activate()

    def use_magic_session(self, magic_session):
        self.magic_session = magic_session
        self._activate()

    def _activate(self):
        # activates this magic_root_session.
        # callbacks wont be ignored, and volume, mute and
        # state will be saved to attributes.
        if not self.magic_app or not self.magic_session:
            self._sav = self._ctl2.QueryInterface(ISimpleAudioVolume)
            self.volume = self._sav.GetMasterVolume()
            self.mute = self._sav.GetMute()
            self.state_id = self._ctl2.GetState()
            new_state = self.AudioSessionState[self.state_id]
            self.state = (new_state, self.state_id)

    def OnSimpleVolumeChanged(self, new_volume, new_mute, event_context):
        log.debug(f"OnSimpleVolumeChanged: {self.app_exe} "
                  f"Vol: ~{new_volume} Mute: {new_mute}")
        # only make callbacks and update internal volume and mute
        # if someone is listening

        # TODO:
        # Now callbacks wich are made by MagicApp or MagicSession
        # get filtered out -> How about passing (new_volume, self_changed)
        # self_changed will be BOLL:
        # True when event_context.contents matches for example:
        # self.magic_app.guid.contents
        # ----
        # or implement a advanced callback witch will pass this?
        # ----
        # or use function.__code__.co_argcount and depending if
        # co_argcount == 1 -> make call only when new
        # co_argcount == 2 -> make call always and pass
        # (new_volume, self_changed)
        # ^
        # | This way obove seems dirty... maybe an explicit flag like:
        # MagicApp(..., always_callback=True) is better?
        # apply same to MagicSession

        # check old volume vs new:
        if self.volume is not None and self.volume != new_volume:
            # self.volume will keep the none state
            # until self._activate()
            self.volume = new_volume

            # get the guid of the 'changer'
            changer_guid = event_context.contents

            # send callbacks, if defined callback exists
            # and if the volume is new:
            if (self.magic_app and
                    changer_guid != self.magic_app.guid.contents):

                if self.magic_app.volume_callback:
                    self.magic_app.volume_callback(new_volume)

            if (self.magic_session and
                    changer_guid != self.magic_session.guid.contents):

                if self.magic_session.volume_callback:
                    self.magic_session.volume_callback(new_volume)

        # check old mute vs new:
        elif self.mute is not None and self.mute != new_mute:
            # self.mute will keep the none state
            # until self._activate()
            self.mute = new_mute

            # get the guid of the 'changer'
            changer_guid = event_context.contents

            # send callbacks, if defined callback exists
            # and if the mute is new:
            if (self.magic_app and
                    changer_guid != self.magic_app.guid.contents and
                    self.magic_app.mute_callback):

                self.magic_app.mute_callback(new_mute)

            if (self.magic_session and
                    changer_guid != self.magic_session.guid.contents and
                    self.magic_session.mute_callback):

                self.magic_session.mute_callback(new_mute)

    def OnStateChanged(self, new_state_id):
        new_state = self.AudioSessionState[new_state_id]

        self.state = (new_state, new_state_id)
        self.state_id = new_state_id

        if new_state_id == 3 and self.magic_manager:
            # if magic_root_session is managed by the magic_manager,
            # wich is the default for use with MagicApp({"app_exe"})
            self.magic_manager.remove_session(self.iid,
                                              self.magic_app)

        # send callbacks, if defined
        if self.magic_app and self.magic_app.state_callback:
            self.magic_app.state_callback(new_state, new_state_id)

        if self.magic_session and self.magic_session.state_callback:
            self.magic_session.state_callback(new_state, new_state_id)

    def _get_app_exe(self):
        self.pid = self._ctl2.GetProcessId()

        if self.pid != 0:
            # try:
            app_exe = psutil.Process(self.pid).name()
            # except psutil.NoSuchProcess:
            # for some reason GetProcessId returned an non existing pid
            # TODO:
            # that should not happen right?
        else:
            # System Sound:
            # self._ctl2.GetDisplayName() returns:
            # @%SystemRoot%\System32\AudioSrv.Dll,-202
            # for system sounds
            # TODO: Double check? or is only self.pid == 0
            # for System Sound
            app_exe = "SndVol.exe"
        return app_exe

    def register_notification(self):
        self._ctl2.RegisterAudioSessionNotification(self)

    def unregister_notification(self):
        self._ctl2.UnregisterAudioSessionNotification(self)
