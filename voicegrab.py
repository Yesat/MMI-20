from mouse._generic import GenericListener as _GenericListener
from mouse._mouse_event import ButtonEvent, MoveEvent, WheelEvent, LEFT, RIGHT, MIDDLE, X, X2, UP, DOWN, DOUBLE
import keyboard
import psutil
import subprocess
import time as _time
import speech_recognition as sr
import os
import platform as _platform
import win32clipboard
from PIL import ImageGrab
from io import BytesIO
from planar import BoundingBox

'''keyboard element'''
# https://github.com/boppreh/keyboard#keyboard.send

''' Mouse elements '''
# https://github.com/boppreh/mouse#mouse.release
if _platform.system() == 'Windows':
    from mouse import _winmouse as _os_mouse
# elif _platform.system() == 'Linux':
#     from mouse import _nixmouse as _os_mouse
else:
    raise OS1Error("Unsupported platform '{}'".format(_platform.system()))


_pressed_events = set()


class _MouseListener(_GenericListener):
    def init(self):
        _os_mouse.init()

    def pre_process_event(self, event):
        if isinstance(event, ButtonEvent):
            if event.event_type in (UP, DOUBLE):
                _pressed_events.discard(event.button)
            else:
                _pressed_events.add(event.button)
        return True

    def listen(self):
        _os_mouse.listen(self.queue)


def drag(start_x, start_y, end_x, end_y, absolute=True, duration=0):
    """
    Holds the left mouse button, moving from start to end position, then
    releases. `absolute` and `duration` are parameters regarding the mouse
    movement.
    """
    if is_pressed():
        release()
    move(start_x, start_y, absolute, 0)
    press()
    move(end_x, end_y, absolute, duration)
    release()


def get_position():
    """ Returns the (x, y) mouse position. """
    return _os_mouse.get_position()


def is_pressed(button=LEFT):
    """ Returns True if the given button is currently pressed. """
    _listener.start_if_necessary()
    return button in _pressed_events


def move(x, y, absolute=True, duration=0):
    """
    Moves the mouse. If `absolute`, to position (x, y), otherwise move relative
    to the current position. If `duration` is non-zero, animates the movement.
    """
    x = int(x)
    y = int(y)

    # Requires an extra system call on Linux, but `move_relative` is measured
    # in millimiters so we would lose precision.
    position_x, position_y = get_position()

    if not absolute:
        x = position_x + x
        y = position_y + y

    if duration:
        start_x = position_x
        start_y = position_y
        dx = x - start_x
        dy = y - start_y

        if dx == 0 and dy == 0:
            _time.sleep(duration)
        else:
            # 120 movements per second.
            # Round and keep float to ensure float division in Python 2
            steps = max(1.0, float(int(duration * 120.0)))
            for i in range(int(steps)+1):
                move(start_x + dx*i/steps, start_y + dy*i/steps)
                _time.sleep(duration/steps)
    else:
        _os_mouse.move_to(x, y)


def press(button=LEFT):
    """ Presses the given button (but doesn't release). """
    _os_mouse.press(button)


def release(button=LEFT):
    """ Releases the given button. """
    _os_mouse.release(button)


_listener = _MouseListener()


''' Check if process is running '''


def checkIfProcessRunning(processName):
    # Check if there is any running process that contains the given name processName.

    # Iterate over the all the running process
    for proc in psutil.process_iter():
        try:
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


'''   Screenshot section '''

def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()


def screenshot(pos_a, pos_b):
    bbox = [min(pos_a[0], pos_b[0]), min(pos_a[1], pos_b[1]), max(pos_a[0], pos_b[0]), max(pos_a[1], pos_b[1])]
    print(bbox)
    image = ImageGrab.grab(bbox)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)


''' Audio recording '''


def get_audio():
    r = sr.Recognizer()

    pos_a = (0, 0)
    pos_b = (0, 0)

    i = 0

    while i < 2:
        print(f'step {i}')
        i += 1
        with sr.Microphone() as source:
            print("I'm listening")
            audio = r.listen(source)
            said = ""
            try:
                said = r.recognize_google(audio)
                print(said)

                if "from" in said:
                    pos_a = get_position()
                    print(f"pos_a = {pos_a}")

                if "there" in said:
                    pos_b = get_position()
                    print(f"pos_b = {pos_b}")   
            except:
                print("Sorry could not understand your command")
                i = i-1
            if pos_b == (0, 0):
                i = 1
                print('something was wrong with second trigger')
            elif pos_a == (0,0):
                i = 0
                print('something was wrong with first trigger')
            print()
    print(pos_a, pos_b)
    return(pos_a,pos_b)


''' Main'''

  
def pushToTalk():

    print("PTT launched")

    # Check if ShareX is already running, if not, launches it
    name = 'ShareX.exe'
    path = "C:\Program Files\ShareX\ShareX.exe"

    if checkIfProcessRunning(name):
        print(f'{name} is running')
    else:
        print(f'{name} isn''t running, lanched')
        p = subprocess.Popen(path)


    # Launches the print sceen shortcut
    # _time.sleep(5)

    print()

    # Record the audio
    _time.sleep(1)
    # print("I'm listening 1")
    pos_a, pos_b = get_audio()

    # print(pos_a)
    # print(pos_a)
    # print(pos_a.y)

    print("drag mouse")
    screenshot(pos_a,pos_b)

    print("PRINT SCREEN")


    print() 

def main():
    hotkey = 'tab + space'

    keyboard.add_hotkey(hotkey, pushToTalk, suppress=True, trigger_on_release=True)


    print("Press tabs + space to start")
    print("Press Esc + space to stop\n")
    keyboard.wait('esc + space')
    print("Code ended")
    keyboard.unhook_all_hotkeys()   

if __name__ == '__main__':
    main()
