from pytest import fixture

from CcReference import CcReference, data


@fixture
def cut():
    yield CcReference()


def test_cc_reference():
    assert len(data) == 133


def test_get_assigned_cc_codes(cut):
    assigned_cc_codes = cut.get_assigned_control_numbers()
    assert len(assigned_cc_codes) > 0
