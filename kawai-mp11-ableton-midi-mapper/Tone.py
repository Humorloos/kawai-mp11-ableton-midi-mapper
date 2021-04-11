import dataclasses as dc
from typing import Dict


@dc.dataclass
class Tone:
    id: int = 0
    map: Dict = dc.field(default_factory=lambda: {})
