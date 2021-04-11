import dataclasses as dc

from Tone import Tone


@dc.dataclass
class Section:
    name: str = ''
    active_tone: Tone = None
