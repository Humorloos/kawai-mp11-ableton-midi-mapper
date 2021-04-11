from typing import Dict

import win32com.client as win32

from CcReference import CcReference
from Section import Section
from SysExInfo import SysExInfo
from cachedproperty import cached_property
from consts import KAWAI_SECTION_NAMES

CONTROL_NUMBER_DICT = CcReference.get_assigned_control_numbers()
SIMPLE_SYS_EX_INFO, PREFIX_SYS_EX_INFO = CcReference.get_cc_sys_ex_info()
MMC_SYS_EX_MAP: Dict[str, SysExInfo] = {info.sys_ex_strings[0]: info for info in SIMPLE_SYS_EX_INFO}
PREFIX_SYS_EX_MAP: Dict[str, SysExInfo] = {sys_ex: info for info in PREFIX_SYS_EX_INFO
                                           for sys_ex in info.sys_ex_strings}
REVERSE_PREFIX_SYS_EX_MAP: Dict[int, SysExInfo] = {info.control_number: info for info in PREFIX_SYS_EX_INFO}
CC_STATUS_OFFSET = 176
PC_STATUS_OFFSET = 192
CA_STATUS_OFFSET = 208
REVERSE_CCS = CcReference.get_reverse_control_numbers()
REVERSE_SECTIONS_DICT = {name: i for i, name in enumerate(KAWAI_SECTION_NAMES)}

mox = None


def int_to_hex(value):
    hex_value = hex(value)[2:].upper()
    if len(hex_value) == 1:
        hex_value = '0' + hex_value
    return hex_value


class EventHandler:

    def __init__(self):
        self.sections = [Section(name=name, active_tone=0) for name in KAWAI_SECTION_NAMES]
        self.active_section = self.sections[1]
        self.out_port_names = {'kawai': '2- KAWAI USB MIDI', 'loopMIDI': 'loopMIDI Port'}
        self.in_port_names = {'kawai': '2- KAWAI USB MIDI', 'loopMIDI': 'loopMIDI Port 1'}

    @cached_property
    def out_ports(self) -> Dict[str, int]:
        # noinspection PyUnresolvedReferences
        return {key: mox.GetOutPortID(port_name) for key, port_name in self.out_port_names.items()}

    @cached_property
    def in_ports(self) -> Dict[str, int]:
        # Ports provided by OnMidiInput() are too small by 1, so we subtract 1 from all IDs
        # noinspection PyUnresolvedReferences
        return {key: mox.GetInPortID(port_name) - 1 for key, port_name in self.in_port_names.items()}

    # noinspection PyPep8Naming
    def OnSysExInput(self, bStrSysEx: str):
        clean_sys_ex = bStrSysEx[:-1]
        if clean_sys_ex in MMC_SYS_EX_MAP.keys():
            sys_ex_info = MMC_SYS_EX_MAP[clean_sys_ex]
            self.output_cc_signal_2_active_track(cc_code=sys_ex_info.control_number, value=127)
        else:
            sys_ex_prefix = bStrSysEx[:29]
            if sys_ex_prefix in PREFIX_SYS_EX_MAP.keys():
                sys_ex_info = PREFIX_SYS_EX_MAP[sys_ex_prefix]
                section_index = sys_ex_info.map[sys_ex_prefix]
                data_size = int(bStrSysEx[27:29], 16)
                data = int(bStrSysEx[27 + 3 * data_size:29 + 3 * data_size], 16)
                if sys_ex_info.name == 'tone':
                    # activate section
                    self.active_section = self.sections[section_index]
                    # activate tone in section
                    if section_index in range(2):
                        self.active_section.active_tone = int(bStrSysEx[30:32] + bStrSysEx[33:35], 16) - (
                                section_index * 12)
                    else:
                        raw_channel = int(bStrSysEx[30:32] + bStrSysEx[33:35], 16) - 24
                        # reserve each 4th tone in sub for return tracks (3, 7, 11, 15 are 12, 13, 14, 15 instead)
                        self.active_section.active_tone = int(raw_channel - raw_channel // 4 + int(
                            not bool((raw_channel + 1) % 4)) * (11 - (2 * ((raw_channel + 1) / 4))))
                    # instead of tone data, always output 127 on respective channel
                    data = 1
                self.output_cc_2_track(track=self.sections[section_index],
                                       cc_code=sys_ex_info.control_number,
                                       value=data * sys_ex_info.scale)

    def activate_active_track(self):
        self.output_cc_signal_2_active_track(cc_code=CONTROL_NUMBER_DICT['tone'], value=127)

    # noinspection PyPep8Naming,PyMethodMayBeStatic
    def OnTerminateMidiInput(self):
        pass

    # noinspection PyPep8Naming,PyUnusedLocal
    def OnMidiInput(self, nTimestamp, port, status, data1, data2):
        if port == self.in_ports['loopMIDI']:
            if data1 in REVERSE_CCS:
                sys_ex_info = REVERSE_PREFIX_SYS_EX_MAP[data1]
                hex_value = int_to_hex(data2 // sys_ex_info.scale)
                some_section_was_updated = False
                affected_tone = status - CC_STATUS_OFFSET
                for i, section in enumerate(self.sections):
                    if section.active_tone == affected_tone:
                        # noinspection PyUnresolvedReferences
                        mox.SendSysExString(f'{sys_ex_info.sys_ex_strings[i]} {hex_value} F7')
                        some_section_was_updated = True
                if not some_section_was_updated:
                    self.active_section.active_tone = affected_tone
                    section_id = REVERSE_SECTIONS_DICT[self.active_section.name]
                    if section_id in range(2):
                        tone_id = affected_tone + section_id * 12
                    else:
                        tone_id = affected_tone + (affected_tone // 3) - (int(affected_tone > 11) * (4 + 3 * (
                                15 - affected_tone)) + int(affected_tone == 15)) + 24
                    # noinspection PyUnresolvedReferences
                    mox.SendSysExString(
                        f'{REVERSE_PREFIX_SYS_EX_MAP[CONTROL_NUMBER_DICT["tone"]].sys_ex_strings[section_id]}'
                        f' 00 {int_to_hex(tone_id)} F7')
                    # noinspection PyUnresolvedReferences
                    mox.SendSysExString(f'{sys_ex_info.sys_ex_strings[section_id]} {hex_value} F7')

    def output_cc_signal_2_active_track(self, cc_code, value):
        self.output_cc_2_track(track=self.active_section, cc_code=cc_code, value=value)

    def output_cc_2_track(self, track: Section, cc_code, value):
        # noinspection PyUnresolvedReferences
        mox.OutputMidiMsg(self.out_ports['loopMIDI'], CC_STATUS_OFFSET + track.active_tone, cc_code,
                          value)


class MidiOxProxy:
    """
    Property and Method Reference
    Informational
    
    --------------------------------------------------------------------------------
    
    GetAppVersion
    
    Returns a string containing the version of MIDI-OX installed in the same directory as the COM interface object. No attachment is required to retrieve this value.
    
    Example:    
    MsgBox "MIDI-OX Version: " & mox.GetAppVersion
    
    
    InstanceCount
    
    This method returns the number of instances of MIDI-OX currently running.
    
    Example:    
    n = mox.InstanceCount
    
    
    InstanceNumber
    
    This method returns the instance number of the attached MIDI-OX.
    
    Example:    
    n = mox.InstanceNumber
    
    
    Script Control
    
    --------------------------------------------------------------------------------
    
    ShouldExitScript
    
    This property is usually queried in a loop, and may be set by MIDI-OX to inform the script that the user has chosen the Exit WScript menu item.
    
    Example:    
    Do While mox.ShouldExitScript = 0
       Msg = mox.GetMidiInputRaw()
       ProcessInput( msg )
    Loop
    
    
    ShutdownAtEnd
    
    This property determines whether MIDI-OX should keep running after the end of the script. By default, the MIDI-OX instance is ended after the script ends. You can change that behavior by using this property.
    
    Example:    
    If vbYes = MsgBox( "Shutdown?", vbYesNo + vbQuestion, "Exit" ) Then
       Mox.ShutdownAtEnd = 1
    Else
       Mox.ShutdownAtEnd = 0
    End If
    
    
    Querying Devices
    
    --------------------------------------------------------------------------------
    
    SysMidiInCount
    
    Returns the number of MIDI input devices installed in Windows.
    
    GetFirstSysMidiInDev
    
    Returns a VB string containing the name of the first MIDI input device defined to Windows.
    
    GetNextSysMidiInDev
    
    Returns a VB string containing the name of the next MIDI input device defined to Windows. Returns a blank string ("") when there are no more devices.
    
    These methods can be used to return all the MIDI input devices installed. They do not require a MIDI-OX attachment (or even that MIDI-OX be running).
    
    Example:    
    str = "Sys MIDI In Devices: " & mox.SysMidiInCount
    strWrk = mox.GetFirstSysMidiInDev
    Do while strWrk <> ""
       str = str & vbCrLf & " " & strWrk
       strWrk = mox.GetNextSysMidiInDev
    Loop
    MsgBox Str
    
    
    SysMidiOutCount
    
    Returns the number of MIDI output devices installed in Windows.
    
    GetFirstSysMidiOutDev
    
    Returns a VB string containing the name of the first MIDI output device defined to Windows.
    
    GetNextSysMidiOutDev
    
    Returns a VB string containing the name of the next MIDI output device defined to Windows. Returns a blank string ("") when there are no more devices.
    
    These methods can be used to return all the MIDI output devices installed. They do not require a MIDI-OX attachment (or even that MIDI-OX be running). Example:    
    str = "Sys MIDI Out Devices: " & mox.SysMidiOutCount
    strWrk = mox.GetFirstSysMidiOutDev
    Do while strWrk <> ""
       str = str & vbCrLf & " " & strWrk
       strWrk = mox.GetNextSysMidiOutDev
    Loop
    MsgBox Str
    
    
    OpenMidiInCount
    
    Returns the number of MIDI input devices opened in the attached MIDI-OX instance.
    
    GetFirstOpenMidiInDev
    
    Returns a VB string containing the name of the first MIDI input device opened in the attached MIDI-OX instance.
    
    GetNextOpenMidiInDev
    
    Returns a VB string containing the name of the next MIDI input device opened in the attached MIDI-OX instance. Returns a blank string ("") when there are no more devices.
    
    These methods can be used to determine which devices are opened by the MIDI-OX instance. They can be compared against the system devices to determine the actual MIDI ID number (it is consecutive).
    
    Example:    
    str = "Open MIDI In Devices: " & mox.OpenMidiInCount
    strWrk = mox.GetFirstOpenMidiInDev
    Do while strWrk <> ""
       str = str & vbCrLf & " " & strWrk
       strWrk = mox.GetNextOpenMidiInDev
    Loop
    MsgBox Str
    
    
    OpenMidiOutCount
    
    Returns the number of MIDI output devices opened in the attached MIDI-OX instance.
    
    GetFirstOpenMidiOutDev
    
    Returns a VB string containing the name of the first MIDI output device opened in the attached MIDI-OX instance.
    
    GetNextOpenMidiOutDev
    
    Returns a VB string containing the name of the first MIDI output device opened in the attached MIDI-OX instance. Returns a blank string ("") when there are no more devices
    
    Example:    
    str = "Open MIDI Out Devices: " & mox.OpenMidiOutCount
    strWrk = mox.GetFirstOpenMidiOutDev
    Do while strWrk <> ""
       str = str & vbCrLf & " " & strWrk
       strWrk = mox.GetNextOpenMidiOutDev
    Loop
    MsgBox Str
    
    
    GetInPortID
    
    Returns an input port number, given a port name.
    
    GetInPortName
    
    Returns an input port name, given a port number.
    
    GetOutPortID
    
    Returns an output port number, given a port name.
    
    GetOutPortName
    
    Returns an output port name, given a port number.
    
    Example:    
    Str = "Sys MIDI In Devices: " & mox.SysMidiInCount & vbCrLf
    StrWrk = mox.GetFirstSysMidiInDev
    Do while strWrk <> ""
       id = mox.GetInPortID( strWrk )
       str = str & vbCrLf & CStr( id ) & ") " & strWrk
       StrWrk = mox.GetNextSysMidiInDev
    Loop
    
    
    Commands
    
    --------------------------------------------------------------------------------
    
    LoadProfile(FilePath)
    
    Opens and applies profile settings to the attached instance.
    
    Example:    
    mox.LoadProfile "C:\Program Files\midiox\myset.ini"
    
    
    LoadDataMap(FilePath)
    
    Opens and turns on the specified .oxm map file.
    
    Example:    
    mox.LoadDataMap "C:\Program Files\midiox\map\transpose up3.oxm"
    
    
    LoadSnapShot(FilePath)
    
    Opens and applies the specified MIDI status info to the MIDI Status View.
    
    Example:    
    mox.LoadSnapShot "C:\Program Files\midiox\snapshot\arden.xms"
    
    
    SendSnapShot
    
    Sends the current MIDI status info out open MIDI ports.
    
    Example:    
    mox.SendSnapShot
    
    
    LoadPatchMap(FilePath)
    
    Opens and applies the specified MIDI Patch-map file into the Patch Mapping view.
    
    Example:    
    mox.LoadPatchMap "C:\Program Files\midiox\map\arden.pmi"
    
    
    System Exclusive
    
    --------------------------------------------------------------------------------
    
    SendSysExFile(FilePath)
    
    This method will send a file containing SysEx out all ports attached by an instance, and which also are mapping the SysEx Port Map object (this is on by default when you open an output port in MIDI-OX). The file is expected to contain encoded binary SysEx data (not ASCII text). If you want to send a file containing SysEx represented as text, open the file and send it via the SendSysExString interface (below).
    
    Example:    
    mox.SendSysExFile "C:\Program Files\midiox\Syx\Sc55.syx"
    
    
    SendSysExString(StrSysEx)
    
    This method will send an ASCII string representing System exclusive data, out all ports attached by an instance, and which also are mapping the SysEx Port Map object (this is on by default when you open an output port in MIDI-OX).
    
    Example:    
    mox.SendSysExString "F0 41 10 42 11 48 00 00 00 1D 10 0B F7"
    
    
    See Also: GetSysExInput (polling) and XXX_SysExInput (event sink) below
    
    MIDI Input and Output
    
    --------------------------------------------------------------------------------
    
    DivertMidiInput
    
    DivertMidiInput is a property that works like a switch: when on MIDI Input is diverted through the script, when off a copy of the MIDI Input is supplied to the script. It can also be queried as if it was a method.
    
    Examples:
    
    mox.DivertMidiInput = 1 ' begin diverting input
    If mox.DivertMidiInput = 1 Then ' query
       MsgBox "Streams are diverted"
    End If
    
    
    FireMidiInput
    
    This property can be set or queried and determines whether MIDI Input should fire a MIDI sink method defined in the script.
    
    Examples:
    
    mox.FireMidiInput = 1 ' begin firing MIDI events
    If mox.FireMidiInput = 0 Then
       MsgBox "Input will not cause trigger"
    End If
    
    
    GetMidiInput
    
    Retrieves MIDI data in the form of a comma delimited string from MIDI-OX. The layout format is as follows: timestamp, port, status, data1, data2. Example: "65123,144,0,77,120". When no input is available an empty string ("") is returned. When System Exclusive data arrives, only the timestamp and a 0xF0 (240) status is received (other data is 0). To retrieve the SysEx string, you need to further call GetSysExInput() before calling GetMidiInput() again.
    
    Example:    
    msgStr = mox.GetMidiInput()
    If msgStr <> "" Then
       A = Split( msgStr, ",", -1, vbTextCompare )
       Tstamp = Int(A(0))
       port = Int(A(1))
       stat = Int(A(2))
       data1 = Int(A(3))
       data2 = Int(A(4))
       If stat = &hF0 Then
          strSysEx = mox.GetSysExInput()
          mox.SendSysExString strSysEx
       End If
    End If
    
    
    GetMidiInputRaw
    
    This method retrieves input in the form of an encoded 32bit value. The timestamp is omitted, and the data is stored with the status in the least significant position (only the lowest 24bits are used to represent 3 MIDI bytes. The format is the same as that supplied to the Windows MME midiOutShortMessage() API. Example: if status=0x90, data1=0x60, data2=0x7F then the value would be encoded as: 0x007F6090. When System Exclusive data arrives, only the 0xF0 (240) status is received (other data is 0). To retrieve the SysEx string, you need to further call GetSysExInput() before calling GetMidiInput() again.
    
    Example:    
    msg = mox.GetMidiInputRaw()
    stat = msg And &h000000F0
    chan = msg And &h0000000F
    msg = msg \ 256 ' pull off stat
    dat1 = msg And &h0000007F
    msg = msg \ 256
    dat2 = msg And &h0000007F
    If stat = &hF0 Then
       strSysEx = mox.GetSysExInput()
       mox.SendSysExString strSysEx
    End If
    
    
    GetSysExInput
    
    This method retrieves System exclusive input after the other two polling methods (GetMidiInput and GetMidiInputRaw), have been advised that SysEx is available (in the form of an 0xF0 satatus value). The SysEx data is formatted into an ASCII string of hex digit bytes separated by spaces. The string is directly compatible with the SendSysExString function. If the SysEx message is longer than 256 data bytes, it is split, and you will receive one or more additional SysEx status bytes.
    
    Example:    
    StrSyx = "F0 43 00 09 F7"
    
    msg = mox.GetMidiInputRaw()
    stat = msg And &h000000F0
    If stat = &hF0 Then
       strSysEx = mox.GetSysExInput()
       mox.SendSysExString strSysEx
    End If
    
    
    OutputMidiMsg(nPort, nStatus, nData1, nData2)
    
    This method sends a MIDI message out all MIDI ports that are mapping the channel. The nStatus parameter is a combination of the MIDI status and channel. Data1 and Data2 values depend on the MIDI message.
    
    Example:    
    mox.OutputMidiMsg -1, 146, 78, 127
    
    
    Connection Point Event Sinks
    
    --------------------------------------------------------------------------------
    
    XXX_MidiInput( timestamp, status, channel, data1, data2 )
    
    This method is called by MIDI-OX whenever data arrives. In order to effect it, you must replace XXX_ with your prefix in both your subroutine definition and your Wscript.CreateObject call.
    
    Example:    
    Set mox = WScript.CreateObject("MIDIOX.MoxScript.1", "Test_")
    
    Sub Test_MidiInput( ts, port, stat, dat1, dat2)
       mox.OutputMidiMsg -1, stat, dat1, dat2
    End Sub
    
    
    XXX_SysExInput()
    
    This method is called by MIDI-OX whenever System Exclusive (SysEx) data arrives. In order to effect it, you must replace XXX_ with your prefix in both your subroutine definition and your Wscript.CreateObject call.
    
    Example:    
    Sub Test_SysExInput( bstrSysEx )
       ' Send it right back
       mox.SendSysExString bstrSysEx
    End Sub
    
    
    XXX_OnTerminateMidiInput()
    
    This event will be triggered by MIDI-OX on the first MIDI message received after the user chooses the Exit Wscript menu item.
    
    Example:    
    Sub Test_OnTerminateMidiInput()
       MsgBox "MIDI Input Termination Received From MIDI-OX"
       mox.FireMidiInput = 0
       mox.DivertMidiInput = 0
    End Sub
    
    
    System Helpers
    
    --------------------------------------------------------------------------------
    
    Sleep( milliseconds )
    
    Pauses the script for at least the specified number of milliseconds. Other processes are allowed to run.
    
    GetSystemTime
    
    Returns the current system time. It represents the number of milliseconds elapsed since Windows was started.
    
    Example:    
    tNow = mox.GetSystemTime
    mox.Sleep(500)
    If tNow + 500 <= mox.GetSystemTime Then
       MsgBox "At least ½ second has passed"
    End If
    """

    def __enter__(self):
        global mox
        mox = win32.DispatchWithEvents("MIDIOX.MoxScript.1", EventHandler)

        mox.DivertMidiInput = 1
        mox.FireMidiInput = 1

        return mox

    def __exit__(self, exc_type, exc_val, exc_tb):
        global mox
        mox.FireMidiInput = 0
        mox.DivertMidiInput = 0

        # Clean up
        mox = None
