from collections.abc import Callable
from typing import TYPE_CHECKING, Any, ParamSpec, TypeVar

import psutil
from typing_extensions import assert_type, override

from pycaw.utils import AudioDevice, AudioSession, AudioUtilities

if TYPE_CHECKING:
    from mypy_extensions import Arg

    from pycaw.api.mmdeviceapi.depend import PROPERTYKEY
    from pycaw.callbacks import (
        AudioEndpointVolumeCallback,
        AudioSessionEvents,
        AudioSessionNotification,
        MMNotificationClient,
    )

    P = ParamSpec("P")
    R = TypeVar("R")

    def callable_contract(callback: Callable[P, R]) -> Callable[P, R]:
        return callback

    class TypedSessionNotification(AudioSessionNotification):
        @override
        def on_session_created(self, new_session: AudioSession) -> None:
            assert_type(new_session, AudioSession)

    session = AudioSession(object())
    output_type: int = 1
    session_notification = AudioSessionNotification()
    session_events = AudioSessionEvents()
    endpoint_events = AudioEndpointVolumeCallback()
    device_events = MMNotificationClient()

    assert_type(AudioUtilities.GetSpeakers(), AudioDevice | None)
    assert_type(AudioUtilities.GetAllSessions(), list[AudioSession])
    assert_type(AudioUtilities.GetProcessSession(123), AudioSession | None)
    assert_type(AudioUtilities.CreateDevice(None), AudioDevice | None)
    assert_type(AudioUtilities.GetAllDevices(), list[AudioDevice])
    assert_type(session.Process, psutil.Process | None)
    assert_type(AudioUtilities.GetEndpointDataFlow("device-id"), str)
    assert_type(AudioUtilities.GetEndpointDataFlow("device-id", 0), str)
    assert_type(AudioUtilities.GetEndpointDataFlow("device-id", 1), int)
    assert_type(AudioUtilities.GetEndpointDataFlow("device-id", output_type), str | int)
    assert_type(
        callable_contract(session_notification.on_session_created),
        Callable[[Arg(AudioSession, "new_session")], None],
    )
    assert_type(
        callable_contract(session_events.on_simple_volume_changed),
        Callable[
            [
                Arg(float, "new_volume"),
                Arg(int, "new_mute"),
                Arg(Any, "event_context"),
            ],
            None,
        ],
    )
    assert_type(
        callable_contract(session_events.on_channel_volume_changed),
        Callable[
            [
                Arg(int, "channel_count"),
                Arg(list[float], "new_channel_volume_array"),
                Arg(int, "changed_channel"),
                Arg(Any, "event_context"),
            ],
            None,
        ],
    )
    assert_type(
        callable_contract(endpoint_events.on_notify),
        Callable[
            [
                Arg(float, "new_volume"),
                Arg(int, "new_mute"),
                Arg(Any, "event_context"),
                Arg(int, "channels"),
                Arg(list[float], "channel_volumes"),
            ],
            None,
        ],
    )
    assert_type(
        callable_contract(device_events.on_default_device_changed),
        Callable[
            [
                Arg(str, "flow"),
                Arg(int, "flow_id"),
                Arg(str, "role"),
                Arg(int, "role_id"),
                Arg(str | None, "default_device_id"),
            ],
            None,
        ],
    )
    assert_type(
        callable_contract(device_events.on_property_value_changed),
        Callable[
            [
                Arg(str, "device_id"),
                Arg(PROPERTYKEY, "property_struct"),
                Arg(Any, "fmtid"),
                Arg(int, "pid"),
            ],
            None,
        ],
    )
