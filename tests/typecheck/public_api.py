from typing import TYPE_CHECKING

import psutil
from typing_extensions import assert_type

from pycaw.utils import AudioDevice, AudioSession, AudioUtilities

if TYPE_CHECKING:
    session = AudioSession(object())
    output_type: int = 1

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
