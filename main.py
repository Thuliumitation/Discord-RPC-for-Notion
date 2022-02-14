from xml.etree.ElementPath import find
import pypresence
import time
import win32gui
import win32process
import pygetwindow as gw
import wmi

c = wmi.WMI()

#Function to get the aplication name from active windows
def get_app_name(hwnd):
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    for p in c.query('SELECT Name FROM Win32_Process WHERE ProcessId = %s' % str(pid)):
        exe = p.Name
        return exe

#Get the handle to the Notion window
def find_notion():
    for i in gw.getAllTitles():
        if get_app_name(win32gui.FindWindow(None, i)) == "Notion.exe":
            hwnd = win32gui.FindWindow(None, i)
            return hwnd
    return None

try:
    client_id = 940875215284609075
    RPC = pypresence.Presence(client_id=client_id)
    RPC.connect()
    print("Connection established!")
    start = int(time.time())
    while True:
        hwnd = find_notion()
        if hwnd is not None:
            RPC.update(state="Organizing life", large_image='notion'
                    , start=start, details=f"Writing: {win32gui.GetWindowText(hwnd)}")
            time.sleep(5)
        else:
            print("You need to run Notion!")
            break
except pypresence.exceptions.DiscordNotFound:
    print("You need to install Discord! If Discord is installed, you need to run it.")
except pypresence.exceptions.ServerError:
    print("Connection to Discord has been closed")
except pypresence.exceptions.InvalidID:
    print("Discord was closed! Closing RPC for Notion")
