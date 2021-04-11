import dataclasses as dc


@dc.dataclass
class Section:
    name: str = ''
    active_tone: int = 0
