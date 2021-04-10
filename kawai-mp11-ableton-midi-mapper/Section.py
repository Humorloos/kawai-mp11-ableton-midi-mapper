import dataclasses as dc


@dc.dataclass
class Section:
    name: str = ''
    is_on: bool = False
    active_tone: int = 0
