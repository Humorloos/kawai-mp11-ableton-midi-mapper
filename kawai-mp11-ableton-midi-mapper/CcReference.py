import os
from typing import List, Dict

import pandas as pd

from SysExInfo import SysExInfo
from consts import KAWAI_SECTION_NAMES

data = pd.read_csv(os.path.join(os.path.dirname(__file__), r'..\resources\cc_midi_reference.csv'), sep=';', skiprows=2)
data[['scale', 'Decimal']] = data[['scale', 'Decimal']].fillna(-1).astype('int32')


class CcReference:
    @staticmethod
    def get_assigned_control_numbers() -> Dict[str, int]:
        return {row['mapping']: row['Decimal'] for i, row in
                data[data['mapping'].notna()][['mapping', 'Decimal']].iterrows()}

    @staticmethod
    def get_reverse_control_numbers():
        return set(data[data['reverse mapping'] == 1]['Decimal'])

    @staticmethod
    def get_cc_sys_ex_info():
        simple_info = [SysExInfo(name=row['mapping'],
                                 control_number=row['Decimal'],
                                 scale=row['scale'],
                                 sys_ex_strings=[row['SysEx PIANO']])
                       for _, row in data[data['SysEx type'] == 'MMC'].iterrows()]
        prefix_info = [SysExInfo(name=row['mapping'],
                                 control_number=row['Decimal'],
                                 scale=row['scale'],
                                 sys_ex_strings=[row[f'SysEx {name}'] for name in KAWAI_SECTION_NAMES])
                       for _, row in data[data['SysEx type'] == 'section'].iterrows()]
        return simple_info, prefix_info
