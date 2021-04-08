import dataclasses as dc


@dc.dataclass
class MidiResponse:
    """
    Data class for storing global and local response codes to midi message
    """
    global_response: int = 0
    local_response: int = 0
