import wmi
import win32con
from win32api import SetCursorPos
from win32gui import (PostMessage, IsWindowVisible, 
                      IsWindowEnabled, EnumWindows,
                      GetWindowRect, IsIconic,
                      ShowWindow, CloseWindow, GetWindowText)
from win32process import GetWindowThreadProcessId
from psutil import process_iter
from time import sleep
from ctypes import (c_char_p, c_ulong, byref, windll)

version = "0.1.6"

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

class Client():
    """
    A Flyff client object with hwnd attribute for handling sending keys to a
    window, and name attribute displaying the logged in character name for
    easy identification (PIDs and hWnds aren't very identifiable).
    """
    def __init__(self, pid):
        self.pid = pid
        self.hwnd = self.get_hwnds()
        self.name = self.get_name()

    def __repr__(self):
        return '<Flyff.Client %s>' % self.pid
    
    def get_hwnds(self):
        """
        Get the hWnd values from the process ID (Flyff only has one hWnd per 
        process, so we just return the one).
        """
        def callback(hwnd, hwnds):
            if IsWindowVisible (hwnd) and IsWindowEnabled(hwnd):
                _, found_pid = GetWindowThreadProcessId(hwnd)
                if found_pid == self.pid:
                    hwnds.append(hwnd)
                    return True
    
        hwnds = []
        EnumWindows(callback, hwnds)
        return hwnds[0]
    
    #def get_name(self):
    #    """
    #    Slightly buggy, but mostly working read of a client's logged in 
    #    character name. Only tested on a single server, so it might not
    #    work at all elsewhere (yet).
    #    The address variable isn't very reliable, as it seems it changes
    #    between patches, and thus very likely different across different clients.
    #    TODO: Find a robust pointer address for this purpose, I guess.
    #    """
    #    OpenProcess = windll.kernel32.OpenProcess
    #    ReadProcessMemory = windll.kernel32.ReadProcessMemory
    #    CloseHandle = windll.kernel32.CloseHandle

    #    PROCESS_ALL_ACCESS = 0x1F0FFF

    #    address = 0x00187E59

    #    buf = c_char_p(b'\x00' * 16)
    #    bufferSize = 16
    #    bytesRead = c_ulong(0)

    #    processHandle = OpenProcess(PROCESS_ALL_ACCESS, False, self.pid)
    #    if ReadProcessMemory(processHandle, 
    #                         address, buf, bufferSize, byref(bytesRead)):
    #        name = buf.value.strip(b'\x06').split(' ')[0]
    #        if len(name) == 0:
    #            name = 'Character not found (not logged in?)'
    #    else:
    #        name = 'Character not found (not logged in?)'
    #    CloseHandle(processHandle)
    #    return name
    
    def get_name(self):
        """
        A certain server has character names in their title, how convenient.
        I'll be using their client to test some things for now.
        """
        return GetWindowText(self.hwnd).split(' ')[0]

class Collector():
    """
    A collector object, with the same attributes as a Client object and some
    coordinates and offsets related to collecting.
    """
    def __init__(self, client):
        self.hwnd = client.hwnd
        self.pid = client.pid
        self.name = client.name
        
        if IsIconic(self.hwnd):
            ShowWindow(self.hwnd, win32con.SW_RESTORE)
        rect = GetWindowRect(self.hwnd)
        self.x, self.y, w, h = rect[0], rect[1], rect[2], rect[3]
        res = ((w - self.x) + (h - self.y))
        CloseWindow(self.hwnd)
        
        # offset_list is still shit, remind me to fix it
        offset_list = [ (1434, (350, 360)), # 800x600
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
                self.offset = offset_tuple[1]
                return
            
    def __repr__(self):
        return '<Flyff.Collector %s>' % self.pid
    

def get_process(n):
    pids = []
    for proc in process_iter():
        if proc.name() == n:
            pids.append(proc.pid)
    return pids

def push_button(hwnd, key):
    """
    Sends a key to a specified hwnd.
    """
    PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
    sleep(0.3) # a small timeout seems to be needed for the key to register
    PostMessage(hwnd, win32con.WM_KEYUP, key, 0)

def click_mouse(hwnd, tx, ty):
    """
    Clicks at x, y in the specified window, relative to the window's upper
    left corner.
    """
    SetCursorPos((tx, ty)) # should be the same for any resolution
    PostMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
    PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)