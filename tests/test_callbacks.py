from ctypes import c_float, cast, create_string_buffer, pointer, sizeof

import pytest
from comtypes import GUID

from pycaw.api.endpointvolume.depend import (
    AUDIO_VOLUME_NOTIFICATION_DATA,
    PAUDIO_VOLUME_NOTIFICATION_DATA,
)
from pycaw.callbacks import AudioEndpointVolumeCallback, AudioSessionEvents


class RecordingEndpointVolumeCallback(AudioEndpointVolumeCallback):
    def on_notify(self, new_volume, new_mute, event_context, channels, channel_volumes):
        self.received = (
            new_volume,
            new_mute,
            event_context,
            channels,
            channel_volumes,
        )


class RecordingSessionEvents(AudioSessionEvents):
    def on_channel_volume_changed(
        self, channel_count, new_channel_volume_array, changed_channel, event_context
    ):
        self.received = (
            channel_count,
            new_channel_volume_array,
            changed_channel,
            event_context,
        )


@pytest.mark.parametrize("channel_count", [1, 2, 8, 9, 12, 17])
def test_endpoint_callback_extracts_all_channel_volumes(channel_count):
    callback = RecordingEndpointVolumeCallback()
    expected_volumes = [index / channel_count for index in range(channel_count)]
    event_context = GUID("{01234567-89AB-CDEF-0123-456789ABCDEF}")

    offset = AUDIO_VOLUME_NOTIFICATION_DATA.afChannelVolumes.offset
    buffer = create_string_buffer(offset + sizeof(c_float) * channel_count)
    notify = cast(buffer, PAUDIO_VOLUME_NOTIFICATION_DATA)
    notify.contents.guidEventContext = event_context
    notify.contents.bMuted = 1
    notify.contents.fMasterVolume = 0.75
    notify.contents.nChannels = channel_count
    channel_array = (c_float * channel_count).from_buffer(buffer, offset)
    channel_array[:] = expected_volumes

    callback.OnNotify(notify)

    new_volume, new_mute, received_context, channels, channel_volumes = (
        callback.received
    )
    assert new_volume == pytest.approx(0.75)
    assert new_mute == 1
    assert received_context.contents == event_context
    assert channels == channel_count
    assert isinstance(channel_volumes, list)
    assert len(channel_volumes) == channel_count
    assert channel_volumes == pytest.approx(expected_volumes)
    assert channel_volumes[-1] == pytest.approx(expected_volumes[-1])


def test_session_callback_copies_reported_channel_volumes_to_list():
    callback = RecordingSessionEvents()
    expected_volumes = [index / 12 for index in range(12)]
    values = (c_float * 13)(*expected_volumes, 1.0)
    event_context = pointer(GUID("{FEDCBA98-7654-3210-FEDC-BA9876543210}"))

    callback.OnChannelVolumeChanged(12, values, 11, event_context)

    channel_count, channel_volumes, changed_channel, received_context = (
        callback.received
    )
    assert channel_count == 12
    assert isinstance(channel_volumes, list)
    assert len(channel_volumes) == channel_count
    assert channel_volumes == pytest.approx(expected_volumes)
    assert changed_channel == 11
    assert received_context is event_context
