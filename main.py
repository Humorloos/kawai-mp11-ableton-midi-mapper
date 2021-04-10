# Press the green button in the gutter to run the script.
import win32con
import win32ui

from MidiOxProxy import MidiOxProxy

if __name__ == '__main__':
    with MidiOxProxy() as _:
        win32ui.MessageBox("Entering Loop... Press OK to end.", "Python", win32con.MB_OK)
