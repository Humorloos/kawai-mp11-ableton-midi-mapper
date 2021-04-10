from pytest import fixture

from MidiOxProxy import MidiOxProxy, EventHandler


@fixture
def cut() -> EventHandler:
    with MidiOxProxy() as mox:
        yield mox._obj_


def test_midi_ox_proxy(cut):
    assert len(cut.sections) == 3
