import time
from tkinter import *
import tkinter as tk
from os.path import expanduser
import os.path
import glob
import os
import json
from win32com.client import Dispatch
import win32com.client
import win32gui
from win32gui import GetWindowText, GetForegroundWindow
import tempfile
from global_hotkeys import *
import pygetwindow as gw
from PIL import Image
from global_hotkeys.keycodes import vk_key_names
import win32con
import tkinter.font as tkFont
import d3dshot
from ahk import AHK
from ahk.window import Window
import win32clipboard
import random

ahk = AHK(executable_path='.\AutoHotkey\AutoHotkey.exe')

speak = Dispatch("SAPI.SpVoice")

Jojoe = False
Enabled = True
version = "1.0"

root = tk.Tk()
resetHotkey = tk.StringVar()
toggleHotkey = tk.StringVar()
borderHotkey = tk.StringVar()
savesPath = StringVar()
currentWorldName = ''
lastCheckedWorld = ''
waitingForQuit = False
speechText = tk.StringVar()
fps = tk.IntVar()
fps.set(30)

mcPid = 0

sid = len(gw.getWindowsWithTitle('SSGMiner v'))
if sid > 0:
    Label(root, text="SeedMiner ID: " + str(sid), fg="green").grid(row=16, sticky=E)

def setDefaults():
    volSlider.set(50)
    resetHotkey.set('end')
    toggleHotkey.set('page_down')
    savesPath.set(expanduser("~") + "\AppData\Roaming\.minecraft\saves")
    speechText.set('Seed')
    borderHotkey.set('delete')
    fps.set(30)
    distSlider.set(16)
    
def enumHandler(mcWin, ctx):
    title = win32gui.GetWindowText(mcWin)
    if title.startswith('Minecraft') and (title[-1].isdigit() or title.endswith('Singleplayer') or title.endswith('Multiplayer (LAN)')):
        style = win32gui.GetWindowLong(mcWin, -16)
        if style == 369623040:
            style = 382664704
            win32gui.SetWindowLong(mcWin, win32con.GWL_STYLE, style)
            if abusePlanar.get():
                win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 0, 290, 1920, 520, 0x0004)
            else:
                win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 500, 270, 920, 540, 0x0004)
        else:
            style &= ~(0x00800000 | 0x00400000 | 0x00040000 | 0x00020000 | 0x00010000 | 0x00800000)
            win32gui.SetWindowLong(mcWin, win32con.GWL_STYLE, style)
            win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 0, 0, 1920, 1080, 0x0004)
        return False
    
def loadConfig():
    global forestIs, beach, desert, plains, tundra, speechText, savesPath, volSlider, Jojoe
    fileConfig = expanduser("~") + "/.ssgResetSettings.json"
    if os.path.isfile(fileConfig):
        try:
            readFile = open(fileConfig, 'r')
            contents = readFile.read()
            readFile.close()
            settings = json.loads(contents)
        except:
            setDefaults()
            saveConfig()
            readFile = open(fileConfig, 'r')
            contents = readFile.read()
            readFile.close()
            settings = json.loads(contents)
        if "savesPath" in settings:
            savesPath.set(settings["savesPath"])
        if "speechText" in settings:
            speechText.set(settings["speechText"])
        if "volume" in settings:
            volSlider.set(settings["volume"])
        else:
            volSlider.set(50)
        if "distance" in settings:
            distSlider.set(settings["distance"])
        else:
            distSlider.set(16)
        if "resetHotkey" in settings:
            resetHotkey.set(settings['resetHotkey'])
        if "toggleHotkey" in settings:
            toggleHotkey.set(settings['toggleHotkey'])
        if "borderHotkey" in settings:
            borderHotkey.set(settings['borderHotkey'])
        if "fps" in settings:
            fps.set(settings['fps'])
    else:
        setDefaults()
        saveConfig()
        
def getMcWin():
    global mcPid
    return ahk.find_window(id=mcPid)
    
def unfocusMc():
    gw.getWindowsWithTitle('SSGMiner v' + version)[0].activate()
    
def selectMC():
    global mcPid
    mcPid = 0
    mcLabel.config(text='Click on Minecraft',fg='red')
    root.update()
    for i in range(20):
        win = ahk.active_window
        try:
            title = win.title.decode(sys.stdout.encoding)
        except:
            title = "aaaaaaaa"
        if title.startswith('Minecraft'):
            if title[-1].isdigit() or title.endswith('Singleplayer') or title.endswith('Multiplayer (LAN)'):
                mcPid = win.id
                mcLabel.config(text='Found Minecraft (' + str(win.pid) + ')',fg='green')
                print('Found MC')
                return
        time.sleep(0.2)
    mcLabel.config(text='Could not find Minecraft',fg='red')

def saveConfig():
    global forestIs, beach, desert, plains, tundra, speecohText, savesPath
    fileConfig = expanduser("~") + "/.ssgResetSettings.json"
    writeFile = open(fileConfig, 'w')
    
    settings = {
    'savesPath':savesPath.get(),
    'speechText':speechText.get(),
    'volume':volSlider.get(),
    'resetHotkey':resetHotkey.get(),
    'toggleHotkey':toggleHotkey.get(),
    'fps':fps.get(),
    'borderHotkey':borderHotkey.get(),
    'distance':distSlider.get()
    }
    
    writeFile.write(json.dumps(settings))
    writeFile.close()
    root.after(1000, saveConfig)
    
def getMostRecentFile(dir):
    try:
        fileList = glob.glob(dir.replace('\\', "/"))
        if not fileList:
            return False
        latest = max(fileList, key=os.path.getctime)
        return latest
    except:
        return False;
    
def isFocused():
    title = GetWindowText(GetForegroundWindow())
    return title.startswith('Minecraft') and (title[-1].isdigit() or title.endswith('Singleplayer') or title.endswith('Multiplayer (LAN)'))

            
fontStyle = tkFont.Font(family="TkDefaultFont", size=9)
volSlider = Scale(root, from_=0, to=100, orient=HORIZONTAL, font=fontStyle)
volSlider.grid(row=10, sticky=W, padx=80)
volLabel = Label(root, text="TTS Volume:", font=fontStyle).grid(row=10, sticky=W, pady=1)
distSlider = Scale(root, from_=0, to=24, orient=HORIZONTAL, font=fontStyle)
distSlider.grid(row=2, sticky=W, padx=0)
loadConfig()
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)

statusLabel = Label(root, text="Running", fg='green', font=fontStyle)
statusLabel.grid(row=1, padx=0, sticky=E)
savesLabel = Label(root, text="Saves folder:", font=fontStyle)
savesLabel.grid(row=6, sticky=W)
savesPathEntry = Entry(root, width=64, exportselection=0, textvariable=savesPath, font=fontStyle).grid(row=7, padx=(5, 0), sticky=W)
speechLabel = Label(root, text="Text to speak:", font=fontStyle).grid(row=8, sticky=W)
speechTextEntry = Entry(root, width=30, exportselection=0, textvariable=speechText, font=fontStyle).grid(row=9, padx=(5, 0), sticky=W)
hotkeyLabel = Label(root, text="Reset World Hotkey:", font=fontStyle)
hotkeyLabel.grid(row=11, sticky=W)
hotkeyEntry = Entry(root, width=20, exportselection=0, textvariable=resetHotkey, font=fontStyle).grid(row=12, padx=5, sticky=W)
toggleHotkeyLabel = Label(root, text="Toggle Resetter Hotkey:", font=fontStyle)
toggleHotkeyLabel.grid(row=13, sticky=W)
toggleHotkeyEntry = Entry(root, width=20, exportselection=0, textvariable=toggleHotkey, font=fontStyle).grid(row=14, padx=5, sticky=W)
warningLabel = Label(root, text="", font=fontStyle)
warningLabel.grid(row=19, sticky=W)
Label(root, text="FPS you record at:", font=fontStyle).grid(row=3, sticky=E)
Radiobutton(root, text="30", padx=7, variable=fps, value=30, font=fontStyle).grid(row=4, sticky=E)
Radiobutton(root, text="60", padx=7, variable=fps, value=60, font=fontStyle).grid(row=5, sticky=E)
Radiobutton(root, text="120", variable=fps, value=120, font=fontStyle).grid(row=6, sticky=E)

borderHotkeyLabel = Label(root, text="Toggle Borderless Hotkey:", font=fontStyle)
borderHotkeyLabel.grid(row=16, sticky=W)
borderHotkeyEntry = Entry(root, width=20, exportselection=0, textvariable=borderHotkey, font=fontStyle).grid(row=17, padx=5, sticky=W)
restartWarning = Label(root, text="Restart after changing binds", font=fontStyle)
restartWarning.grid(row=18, sticky=W)

Button(root, text="Assign MC Window", command=selectMC).grid(row=18, sticky=E, padx=5)
mcLabel = Label(root, text="No Minecraft window found", fg="red", font=fontStyle)
mcLabel.grid(row=17, sticky=E)

Label(root, text="Distance from optimal:", font=fontStyle).grid(row=1, sticky=W, pady=1)
Label(root, text="(Centered at -239, 251)", font=fontStyle).grid(row=3, sticky=W, pady=1)

root.resizable(False, False)
root.title("SSGMiner v" + version)

def scanForMc():
    global mcPid
    window = ahk.find_window(title=b'Minecraft* 1.16.1')
    if not window:
        window = ahk.find_window(title=b'Minecraft* 1.16.1 - Singleplayer')
        if not window:
            window = ahk.find_window(title=b'Minecraft* 1.16.1 - Multiplayer (LAN)')
            if not window:
                print('Can''t find MC F')
                return
    mcPid = window.id
    mcLabel.config(text='Found Minecraft (' + str(window.pid) + ')',fg='green')

def canCheck():
    global currentWorldName, waitingForQuit, lastCheckedWorld
    currentWorld = getMostRecentFile(savesPath.get() + "/*")
    # adv folder does not exist during early world creation
    if currentWorld == False:
        savesLabel.config(text="Saves folder: (currently invalid)", fg='red')
        return False
    if savesLabel['text'] == "Saves folder: (currently invalid)":
        savesLabel.config(text="Saves folder:", fg='black')
    if not (os.path.isdir(currentWorld + "/advancements")):
        return False
    if waitingForQuit == True:
        try: 
            lockFile = open(currentWorld + "/session.lock", "r")
            # fails with error 13 (permission denied) if world is running
            lockFile.read()
            waitingForQuit = False
            makeWorld()
        except IOError as e:
            return False
    advCreation = os.stat(currentWorld + "/advancements").st_ctime
    timeElapsed = time.time() - advCreation;
    return timeElapsed < 5 and lastCheckedWorld != os.path.basename(currentWorld) and waitingForQuit == False

def reportSeed():
    global speechText, mcPid
    win = getMcWin()
    speak.Volume = volSlider.get()
    if random.random() > 0.9:
        speak.Rate = 1
    if random.random() > 0.99:
        speak.rate = -10
    speak.Speak(speechText.get())
    return
    
def resetRun(pausedAlready = False):
    global waitingForQuit
    waitingForQuit = True
    win = getMcWin()
    if not pausedAlready:
        time.sleep(0.1)
        win.send('{escape}')
    time.sleep(0.05)
    win.send('{shift Down}{tab}{shift Up}')
    time.sleep(0.05)
    win.send('{enter}')

def checkBiome():
    global currentWorldName, waitingForQuit, lastCheckedWorld, currentWorldLabel
    currentWorld = getMostRecentFile(savesPath.get() + "/*")
    if currentWorld == False:
        print('cant check biome, no world')
        return
    lastCheckedWorld = os.path.basename(currentWorld)

    win = getMcWin()
    win.send('{f3 Down}')
    time.sleep(0.01)
    win.send('c')
    time.sleep(0.01)
    win.send('d')
    time.sleep(0.01)
    win.send('{f3 Up}')
    time.sleep(0.01)
    win.send('{escape}')
    time.sleep(0.1)
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData().split()
    win32clipboard.CloseClipboard()
    base = (-239, 251)
    dist = distSlider.get()
    bound1 = (base[0] - dist, base[1] - dist)
    bound2 = (base[0] + dist, base[1] + dist)
    x = float(data[6])
    y = float(data[7])
    z = float(data[8])
    if x >= bound1[0] and x <= bound2[0] and z >= bound1[1] and z <= bound2[1]:
        reportSeed()
        return
    resetRun(True)
            
def waitForColours():
    print('check for color')
    d = d3dshot.create(capture_output="pil")
    win = getMcWin()
    rect = win.rect
    rect = (rect[0] + 50, rect[1] + 200, rect[0] + 52, rect[1] + 202)
    img = d.screenshot(region=rect)
    color = img.getpixel((1, 1))
    print(color)
    startTime = time.time()
    while (time.time() - startTime) < 0.5 and not (color[0] > 13 and color[0 < 18] and color[1] > 9 and color[1] < 15 and color[2] > 5 and color[2] < 10):
        img = d.screenshot(region=rect)
        color = img.getpixel((1, 1))
        time.sleep(0.05)

def waitForWorlds():
    waitForColours()
    time.sleep(0.1)

def makeWorld():
    delay = 0.08
    if fps.get() == 60:
        delay /= 2
    elif fps.get() == 120:
        delay /= 4
    win = getMcWin()
    time.sleep(0.1)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{enter}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{enter}')
    time.sleep(delay)
    #6 tabs
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    if True:
        win.send('{enter}')
        time.sleep(delay)
        win.send('{enter}')
        time.sleep(delay)
        win.send('{enter}')
        time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{enter}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('2483313382402348964')
    time.sleep(0.5)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{enter}')
    
def hotkeyReset():
    currentWorld = getMostRecentFile(savesPath.get() + "/*")
    if currentWorld == False:
        print('no world found')
        return;
    win = getMcWin()
    if win.active:
        gw.getWindowsWithTitle('SSGMiner v' + version)[0].activate()
        time.sleep(0.05)
    try:
        lockFile = open(currentWorld + "/session.lock", "r")
        # fails with error 13 (permission denied) if world is running
        lockFile.read()
        makeWorld()
    except IOError as e:
        if e.args[0] == 13:
            resetRun()
        else:
            print('Some generic error happened when checking world hotkey reset idk ' + e.args[1])
    
def toggleEnabled():
    global Enabled
    if Enabled:
        currentWorldName = ''
        lastCheckedWorld = ''
        waitingForQuit = False
        Enabled = False
        statusLabel.config(text='Paused', fg='red')
    else:
        Enabled = True
        statusLabel.config(text='Running', fg='green')

def toggleBorder():
    try:
        win32gui.EnumWindows(enumHandler, None)
    except:
        return

if resetHotkey.get() != '':
    bindings = [
        [[resetHotkey.get()], None, hotkeyReset],
        [[toggleHotkey.get()], None, toggleEnabled],
        [[borderHotkey.get()], None, toggleBorder]
    ]
    try:
        register_hotkeys(bindings)
    except:
        print('Error, invalid hotkey')
    start_checking_hotkeys()
    
def checkHotkeys():
    if not resetHotkey.get().lower() in vk_key_names:
        hotkeyLabel.config(text='Reset Hotkey (INVALID):', fg='red')
        restartWarning.config(fg='red')
    else:
        hotkeyLabel.config(text='Reset Hotkey:', fg='black')
    if not toggleHotkey.get().lower() in vk_key_names:
        toggleHotkeyLabel.config(text='Toggle Resetter Hotkey (INVALID):', fg='red')
        restartWarning.config(fg='red')
    else:
        toggleHotkeyLabel.config(text='Toggle Resetter Hotkey:', fg='black')
    if not borderHotkey.get().lower() in vk_key_names:
        borderHotkeyLabel.config(text='Toggle Borderless Hotkey (INVALID):', fg='red')
        restartWarning.config(fg='red')
    else:
        borderHotkeyLabel.config(text='Toggle Borderless Hotkey:', fg='black')
    
    root.after(500, checkHotkeys)

scanForMc()
    
def mainLoop():
    if Enabled and canCheck():
        checkBiome()
    root.after(50, mainLoop)
root.after(0, mainLoop)
root.after(0, checkHotkeys)
root.after(1000, saveConfig)
root.geometry('310x415')
root.mainloop()
stop_checking_hotkeys()
