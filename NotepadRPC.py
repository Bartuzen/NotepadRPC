print("Starting NotepadRPC.py 0.2")
import os
import sys
import time
import ctypes
import calendar
from datetime import datetime
lib = []
try: import pypresence
except ModuleNotFoundError: lib.append("pypresence")
try: from ruamel.yaml import YAML
except ModuleNotFoundError: lib.append("ruamel.yaml")
if len(lib) > 0:
    print("Couldn't find required libraries. Trying to install them...")
    for each in lib:
        response = os.system('{} -m pip install -U '.format(sys.executable)+each+" -q")
        if response != 0:
            print("Couldn't install "+each)
            time.sleep(5)
            sys.exit()
    print("Successfully installed libraries.")
    try:
        import pypresence
        from ruamel.yaml import YAML
    except: pass
yaml = YAML()
if os.path.exists("config.yml"):
    if not os.path.isfile("config.yml"):
        print("Config.yml is not a file.")
        time.sleep(5)
        sys.exit()
else:
    open("config.yml","w+")
    c = True
try:
    with open("config.yml","r", encoding='utf8') as stream:
        config = yaml.load(stream.read()) or {}
except Exception as e:
    print("Encountered an error about config: "+str(e))
    time.sleep(5)
    sys.exit()
def gen_yaml(key, value):
    global config
    if key not in config:
        config[key] = value
        return True
re = []
re.append(gen_yaml("clientid",123456789))
re.append(gen_yaml("sleep",0))
if True in re:
    yaml.dump(config,open("config.yml","w", encoding='utf8'))
if "c" in globals():
   print("Generated config. Edit it and start program again.")
   time.sleep(5)
   sys.exit()
try:
    if not os.path.exists("extensions.yml"):
        open("extensions.yml","w+")
    with open("extensions.yml","r", encoding='utf8') as stream:
        ext = yaml.load(stream.read()) or {}
except:
    print("Couldn't open extensions.yml. Ignoring it.")
    ext = {}
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
def get_np():
    titles = []
    def foreach_window(hwnd, lParam):
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
rpc = pypresence.Presence(config["clientid"])
if get_np() != None:
    rpc.connect()
    status = 1
else:
    status = 0
while True:
    try:
        if status == 1:
            if get_np() == None:
                rpc.close()
                status = 0
                if "file" in globals():
                    del file
                continue
            if "file" not in globals() or file != get_np():
                t = calendar.timegm(datetime.utcnow().utctimetuple())
                file = get_np()
            if file != None and not file.startswith("new") and not file.startswith("*new") and (os.path.isfile(file) or os.path.isfile(file[1:])):
                try: size = os.path.getsize(file) if file[0] != "*" else os.path.getsize(file[1:])
                except: size_t = None
                else:
                    if size <= 1024:
                        size_t = str(size)+" bytes"
                    elif size < 1024*1024:
                        size_t = str(size/1024).split(".")[0]+("."+str(size/1024).split(".")[1][:2] if str(size/1024).split(".")[1] != "0" else "")+" kilobytes"
                    elif size < 1024*1024*1024:
                        size_t = str(size/(1024*1024)).split(".")[0]+("."+str(size/(1024*1024)).split(".")[1][:2] if str(size/(1024*1024)).split(".")[1] != "0" else "")+" megabytes"
                    else:
                        size_t = str(size/(1024*1024*1024)).split(".")[0]+("."+str(size/(1024*1024*1024)).split(".")[1][:2] if str(size/(1024*1024*1024)).split(".")[1] != "0" else "")+" gigabytes"
            else:
                size_t = None
            rpc.update(details="Editing "+file.split("\\")[-1] if file != None else "Idling",
                       state=size_t,
                       large_image="image_large",
                       large_text="Notepad++",
                       small_image=("image_"+file.split("\\")[-1].split(".")[-1]) if file != None and "." in file.split("\\")[-1] else None,
                       small_text=ext[file.split("\\")[-1].split(".")[-1]] if file != None and "." in file.split("\\")[-1] and file.split("\\")[-1].split(".")[-1] in ext else None,
                       start=t if file != None else None)
        elif get_np() != None:
            rpc = pypresence.Presence(config["clientid"])
            rpc.connect()
            status = 1
    except pypresence.exceptions.InvalidID:
        print("Invalid ClientID")
        time.sleep(5)
        sys.exit()
    except Exception as e:
        print("An error has occured: "+str(e))
    else:
        if "start" not in globals():
            start = 1
            print("NotepadRPC.py has started successfully!")
    time.sleep(float(str(config["sleep"])))