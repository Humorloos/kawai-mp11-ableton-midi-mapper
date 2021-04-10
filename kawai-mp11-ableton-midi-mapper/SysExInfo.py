import dataclasses as dc
from typing import List

from cachedproperty import cached_property


@dc.dataclass
class SysExInfo:
    name: str = ''
    control_number: int = 0
    scale: int = 0
    sys_ex_strings: List[str] = None

    @cached_property
    def map(self):
        return {sys_ex: i for i, sys_ex in enumerate(self.sys_ex_strings)}

    @cached_property
    def reverse_map(self):
        return {i: sys_ex for i, sys_ex in enumerate(self.sys_ex_strings)}
