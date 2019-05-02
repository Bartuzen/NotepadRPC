VERSION = "0.4"
print("Starting NotepadRPC.py " + VERSION)
import os
import sys
import time
import ctypes
import calendar
from datetime import datetime
lib = []
try:
    import pypresence
except ModuleNotFoundError:
    lib.append("pypresence")
try:
    from ruamel.yaml import YAML
except ModuleNotFoundError:
    lib.append("ruamel.yaml")
if len(lib) > 0:
    print("Couldn't find required libraries. Trying to install them...")
    for each in lib:
        response = os.system('{} -m pip install -U '.format(sys.executable)+each+" -q")
        if response != 0:
            sys.exit("Couldn't install " + each)
    print("Successfully installed libraries.")
    import pypresence
    from ruamel.yaml import YAML
yaml = YAML()
if os.path.exists("config.yml"):
    if not os.path.isfile("config.yml"):
        sys.exit("Config.yml is not a file.")
else:
    open("config.yml", "w+")
    print("Generated config.")
with open("config.yml", "r", encoding='utf8') as stream:
    config = yaml.load(stream.read()) or {}


def gen_yaml(key, value):
    global config
    if key not in config:
        config[key] = value
        return True


re = list()
re.append(gen_yaml("clientId", 529306098646122516))
re.append(gen_yaml("sleep", 1))
re.append(gen_yaml("fileSwitchResetsTimer", True))
if True in re:
    yaml.dump(config, open("config.yml", "w", encoding='utf8'))

if os.path.exists("extensions.yml"):
    if not os.path.isfile("extensions.yml"):
        sys.exit("Extensions.yml is not a file.")
else:
    open("extensions.yml", "w+")
    print("Generated extensions file.")
with open("extensions.yml", "r", encoding='utf8') as stream:
    ext = yaml.load(stream.read()) or {}

EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible


def get_np():
    titles = []

    def foreach_window(hwnd):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buff, length + 1)
            titles.append(buff.value)
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    for each in titles:
        if each[-12:] == " - Notepad++":
            return each[:-12]


def float_format(number):
    number = str(number)
    if number.split(".")[1] == "0":
        return number.split(".")[0]
    number_split = number.split(".")
    return number_split[0] + "." + number_split[1][:2]


connected = False
rpc = pypresence.Presence(config["clientId"])


def presence():
    global rpc
    global config
    global started
    global connected
    file = get_np()
    if file is not None:
        if not connected:
            rpc.connect()
            connected = True
            started = calendar.timegm(datetime.utcnow().utctimetuple())
        if file.startswith("*"):
            file = file[1:]
        name = file.split("\\")[-1]
        if not name.startswith("new ") and "." in name:
            extension = name.split(".")[-1]
        else:
            extension = None
        if config["fileSwitchResetsTimer"]:
            global old_file
            if "old_file" in globals() and old_file != file:
                started = calendar.timegm(datetime.utcnow().utctimetuple())
            old_file = file
        try:
            if name.startswith("*new ") or name.startswith("new "):
                raise Exception
            size = os.path.getsize(file) if file[0] != "*" else os.path.getsize(file[1:])
            if size <= 1024:
                size_text = str(size) + " bytes"
            elif size < 1024 * 1024:
                size_text = float_format(size/1024) + " kilobytes"
            elif size < 1024 * 1024 * 1024:
                size_text = float_format(size/(1024*1024)) + " megabytes"
            elif size < 1024 * 1024 * 1024:
                size_text = float_format(size/(1024*1024*1024)) + " gigabytes"
        except:
            size_text = None
        rpc.update(details="Editing " + name,
                   state=size_text,
                   large_image="image_large",
                   small_image="image_" + extension if extension is not None else None,
                   small_text=ext[extension] if extension is not None and extension in ext else None,
                   start=started)
    elif connected:
        rpc.close()
        connected = False


try:
    presence()
except pypresence.exceptions.InvalidID:
    print()
    sys.exit("Invalid clientId")
else:
    print("NotepadRPC.py has started successfully!")
while True:
    time.sleep(float(config["sleep"]))
    presence()
