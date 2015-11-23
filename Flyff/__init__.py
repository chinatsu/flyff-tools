import wmi
import win32con
from win32api import SetCursorPos
from win32gui import (PostMessage, IsWindowVisible, 
                      IsWindowEnabled, EnumWindows)
from win32process import GetWindowThreadProcessId
from time import sleep

version = "0.1.2"

keys = {"F1": win32con.VK_F1,
        "F2": win32con.VK_F2,
        "F3": win32con.VK_F3,
        "F4": win32con.VK_F4,
        "F5": win32con.VK_F5,
        "F6": win32con.VK_F6,
        "F7": win32con.VK_F7,
        "F8": win32con.VK_F8,
        "F9": win32con.VK_F9,
        }

def get_process(n):
    """
	Get all processes with the specified name and return their process IDs in
	a list.
	"""
    c = wmi.WMI()
    pids = []
    print "Mapping processes..."
    for process in c.Win32_Process():
        if process.name == n:
            pids.append(process.ProcessId)
    return pids

def get_hwnds(pid):
	"""
	Get the hWnd values from the process ID (Flyff only has one hWnd per 
	process).
	"""
    def callback(hwnd, hwnds):
        if IsWindowVisible (hwnd) and IsWindowEnabled(hwnd):
            _, found_pid = GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                hwnds.append(hwnd)
                return True
    
    hwnds = []
    EnumWindows(callback, hwnds)
    return hwnds[0]

def push_button(hwnd, key):
    """
	Sends a key to a specified hwnd.
	"""
    PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
    sleep(0.3)
    PostMessage(hwnd, win32con.WM_KEYUP, key, 0)

def click_mouse(hwnd, tx, ty):
    """
	Clicks at x, y in the specified window, relative to the window's upper
	left corner.
    """
    SetCursorPos((tx, ty)) # should be the same for any resolution
    PostMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
    PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)

def get_offset(res):
    """Get coordinate offsets for the confirmation button when using batteries.
    It uses a window's height + width, named res for lookup-purposes, and
    a match has to be within 5 of what I get on Windows 7. W10 seems to get 
    slightly different values.
	These values may be broken for certain clients, and needs revising.
    """
    offset_list = [(1434, (350, 360)), # 800x600
                   (1826, (465, 445)), # 1024x768
                   (2034, (590, 420)), # 1280x720 
                   (2082, (590, 445)), # 1280x768 
                   (2114, (590, 460)), # 1280x800 
                   (2162, (630, 445)), # 1360x768 
                   (2338, (590, 570)), # 1280x1024 
                   (2390, (670, 510)), # 1440x900
                   (2484, (650, 585)), # 1400x1050 
                   (2534, (750, 510)), # 1600x900 
                   (2698, (750, 595)), # 1600x1200 
                   (2764, (790, 585)), # 1680x1050 
                   (3018, (910, 595)) # 1920x1080 
                   ]
    for offset_tuple in offset_list:
        if offset_tuple[0]-5 <= res <= offset_tuple[0]+5:
            offset = offset_tuple[1]
            return offset