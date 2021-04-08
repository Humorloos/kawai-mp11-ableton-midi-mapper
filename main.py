import re
from typing import Dict

import win32com.client as win32
import win32con
import win32ui

from MidiResponse import MidiResponse
from cachedproperty import cached_property

SYS_EX_TOGGLE_MAP = {'F0 7F 00 06 06 F7 ': [3, False]}  # RECORD

SYS_EX_SIMPLE_MAP = {'F0 7F 00 06 02 F7 ': MidiResponse(0, 9),  # PLAY
                     'F0 7F 00 06 04 F7 ': MidiResponse(14, 0),  # FAST FORWARD
                     'F0 7F 00 06 05 F7 ': MidiResponse(15, 0),  # REWIND
                     'F0 7F 00 06 08 F7 ': MidiResponse(21, 20)}  # RECORD PAUSE
SYS_EX_FADER_MAP = {'F0 40 00 10 00 12 40 01 70 01 ': 22,
                    'F0 40 00 10 00 12 40 04 38 01 ': 23}  # some volume fader (not sure which)

banks_regex = r'F0 40 00 10 00 12 40 00 49 01 0(.) F7'  # banks on/off
bank_regex = r'F0 40 00 10 00 12 40 00 19 02 0([01]) 0([0-7]) F7'  # selecting bank A1 - B8
bank_control_number = 110
cc_status_offset = 176
pc_status_offset = 192
ca_status_offset = 208


class EventHandler:

    def __init__(self):
        self.banks_active = False
        self.active_bank: int = 0
        self.ports = {'kawai': '2- KAWAI USB MIDI', 'loopMIDI': 'loopMIDI Port'}

    @cached_property
    def out_ports(self) -> Dict[str, int]:
        return {key: mox.GetOutPortID(port_name) for key, port_name in self.ports.items()}

    @cached_property
    def in_ports(self) -> Dict[str, int]:
        return {key: mox.GetInPortID(port_name) - 1 for key, port_name in self.ports.items()}

    # noinspection PyPep8Naming
    def OnSysExInput(self, bStrSysEx: str):
        if bStrSysEx in SYS_EX_TOGGLE_MAP.keys():
            self.output_cc_signal_2_loopMIDI(cc_code=SYS_EX_TOGGLE_MAP[bStrSysEx][0],
                                             value=SYS_EX_TOGGLE_MAP[bStrSysEx][1] * 127)
            SYS_EX_TOGGLE_MAP[bStrSysEx][1] = not SYS_EX_TOGGLE_MAP[bStrSysEx][1]
        elif bStrSysEx in SYS_EX_SIMPLE_MAP.keys():
            if self.banks_active:
                self.output_cc_signal_2_loopMIDI(cc_code=SYS_EX_SIMPLE_MAP[bStrSysEx].local_response, value=127)
            else:
                self.output_cc_signal_2_loopMIDI(cc_code=SYS_EX_SIMPLE_MAP[bStrSysEx].global_response, value=127)
        elif bStrSysEx[:30] in SYS_EX_FADER_MAP.keys():
            self.output_cc_signal_2_loopMIDI(cc_code=SYS_EX_FADER_MAP[bStrSysEx[:30]], value=int(bStrSysEx[30:32], 16))
        else:
            banks_match = re.match(banks_regex, bStrSysEx)
            if banks_match:
                if banks_match.group(1) == '1':
                    self.banks_active = True
                else:
                    self.banks_active = False
                    self.active_bank = 0
                return
            bank_match = re.match(bank_regex, bStrSysEx)
            if bank_match:
                self.active_bank = 8 * int(bank_match.group(1)) + int(bank_match.group(2))
                self.output_cc_signal_2_loopMIDI(cc_code=bank_control_number, value=127)

        # mox.SendSysExString(bStrSysEx)

    # noinspection PyPep8Naming,PyMethodMayBeStatic
    def OnTerminateMidiInput(self):
        pass

    # noinspection PyPep8Naming,PyUnusedLocal
    def OnMidiInput(self, nTimestamp, port, status, data1, data2):
        if port == self.in_ports['kawai']:
            if data1 != 64:
                if status in range(cc_status_offset, cc_status_offset + 16):
                    mox.OutputMidiMsg(self.out_ports['loopMIDI'], status + self.active_bank, data1, data2)
                else:
                    self.output_cc_signal_2_loopMIDI(cc_code=data1, value=data2)
        elif port == self.in_ports['loopMIDI']:
            pass

    # noinspection PyPep8Naming
    def output_cc_signal_2_loopMIDI(self, cc_code, value):
        mox.OutputMidiMsg(self.out_ports['loopMIDI'], cc_status_offset + self.active_bank, cc_code, value)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    mox = win32.DispatchWithEvents("MIDIOX.MoxScript.1", EventHandler)

    mox.DivertMidiInput = 1
    mox.FireMidiInput = 1

    win32ui.MessageBox("Entering Loop... Press OK to end.", "Python", win32con.MB_OK)

    mox.FireMidiInput = 0
    mox.DivertMidiInput = 0

    # Clean up
    mox = None
