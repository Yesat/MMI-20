from mouse._generic import GenericListener as _GenericListener
from mouse._mouse_event import ButtonEvent, MoveEvent, WheelEvent, LEFT, RIGHT, MIDDLE, X, X2, UP, DOWN, DOUBLE

from PIL import ImageGrab
from io import BytesIO
from gtts import gTTS

import keyboard
import psutil
import subprocess
import time as _time
import speech_recognition as sr
import os
import platform as _platform
import win32clipboard
from playsound import playsound


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
    image = ImageGrab.grab(bbox)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)
    print('screenshot saved to the clipboard')

    
    
    
''' Acknowledgment sounds'''

def createSounds(filename, text):
    tts = gTTS(text=text, lang="en")
    tts.save(filename)

    


''' Audio recording '''
     
def get_audio():
    r = sr.Recognizer()

    pos_a = (0, 0)   
    pos_b = (0, 0)

    i = 0
    while i < 2:
        commands = ['from here', 'to there']
        i += 1
        print(f'step {i}')
        
        with sr.Microphone() as source:
            print("I'm listening")
            audio = r.listen(source)
            said = ""
            try:
                said = r.recognize_google(audio)
                print(said)

                if "from here" in said and i==1:
                    pos_a = get_position()
                    print(f"pos_a = {pos_a}")
                    playsound("posA.mp3")

                if "there" in said and i==2:
                    pos_b = get_position()
                    print(f"pos_b = {pos_b}")   
                    playsound("posB.mp3")
            except:
                playsound("notUnderstood.mp3")
                print("Sorry could not understand your command")
                i = i-1
            if pos_a == (0,0):
                i = 0
                playsound("repeatA.mp3")
                print('Please say "from here" for the software to place the first point of the screenshot')
            elif pos_b == (0, 0) and i == 2:
                i = 1
                playsound("repeatB.mp3")
                print('Please say "to there" for the software to place the end point of your screenshot')
            print() 
    return(pos_a,pos_b)


''' Main'''
 

def pushToTalk():

    print("\n\nPTT launched")
    
#     Listening acknowledgment for the user
    playsound("listeningAck.mp3")

    
#     Record the audio
    pos_a, pos_b = get_audio()


    screenshot(pos_a,pos_b)

    print() 
    playsound("saved.mp3")

    playsound("exit.mp3")

    
def main():
    
    # quick check and folder creation if it was removed
    ready_file = os.path.join('.', 'ready.mp3')
    listeningAck_file = os.path.join('.', 'listeningAck.mp3')
    posA_file = os.path.join('.', 'posA.mp3')
    posB_file = os.path.join('.', 'posB.mp3')
    repeatA_file = os.path.join('.', 'repeatA.mp3')
    repeatB_file = os.path.join('.', 'repeatB.mp3')
    notUnderstood_file = os.path.join('.', 'notUnderstood.mp3')
    saved_file = os.path.join('.', 'saved.mp3')
    end_file = os.path.join('.', 'end.mp3')
    exit_file = os.path.join('.', 'exit.mp3')

    # Check if the file already exists
    if not os.path.isfile(ready_file):
        print("Creating audio file ready")
        createSounds("ready.mp3", "Hello there")
        
    if not os.path.isfile(listeningAck_file):
        print("Creating audio file listeningAck")
        createSounds("listeningAck.mp3", "I am listening")
        
    if not os.path.isfile(posA_file):
        print("Creating audio file posA")
        createSounds("posA.mp3", "First position saved")
        
    if not os.path.isfile(posB_file):
        print("Creating audio file posB")
        createSounds("posB.mp3", "Second position saved")
        
    if not os.path.isfile(repeatA_file):
        print("Creating audio file repeatA")
        createSounds("repeatA.mp3", "Please say 'from here' for the first command")
        
    if not os.path.isfile(repeatB_file):
        print("Creating audio file repeatB")
        createSounds("repeatB.mp3", "Please say 'to there' for the second command.")
        
    if not os.path.isfile(notUnderstood_file):
        print("Creating audio file notUnderstood")
        createSounds("notUnderstood.mp3", "Sorry did not understand your command")
        
    if not os.path.isfile(saved_file):
        print("Creating audio file saved")
        createSounds("saved.mp3", "Your screenshot has been copied")
        
    if not os.path.isfile(end_file):
        print("Creating audio file end")
        createSounds("end.mp3", "Goodbye!")
        
    if not os.path.isfile(exit_file):
        print("Creating audio file exit")
        createSounds("exit.mp3", "You can take another screenshot by pressing the key above tabulation or quit by pressing escape")
        

    
    
    hotkey = '§'

    keyboard.add_hotkey(hotkey, pushToTalk, suppress=True, trigger_on_release=True)

    print('To take a screenshot: (of any size)\n\n - press the § key of your keyboard (under the Esc key)\n\n - Wait fo the phrase "I am listening"\n\n - then say "Take a screenchot from here", here being the position of the mouse\'s pointer and where you want the screenshot to start\n\n - wait for the acknowledgement "The first position is saved",\n\n - then say "to there", there being the position of the mouse\'s pointer and where you want the screenshot to stop.\n\n The screenshot is then saved into your clipboard and can be pasted anywhere you want.\n')

    print("Press § to start")
    print("Press Esc to stop\n")
    playsound("ready.mp3")
    keyboard.wait('esc')
    
    playsound("end.mp3")
    print("Code ended")
    
    keyboard.unhook_all_hotkeys()   

if __name__ == '__main__':
    main()
