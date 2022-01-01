import win32gui, win32con, win32api, ctypes, time
import pandas as pd
from pywinauto import clipboard
from difflib import SequenceMatcher

room = win32gui.FindWindow(None, "바위게")
inBox = win32gui.FindWindowEx(room, None , "RICHEDIT50W" , None)  # 채팅창의 메세지 입력창
with open("QnA.txt", "r", encoding="UTF-8") as f:
    lines = f.readlines()

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con

def kakao_sendtext(inText):
    win32api.SendMessage(inBox, win32con.WM_SETTEXT, 0, inText) # 채팅창 입력
    win32api.PostMessage(inBox, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32api.PostMessage(inBox, win32con.WM_KEYUP, win32con.VK_RETURN, 0) # 엔터키

def get_chat():
    hwndListControl = win32gui.FindWindowEx(room, None, "EVA_VH_ListControl_Dblclk", None)
    PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    get_text = clipboard.GetData()
    return get_text

def PostKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):
        ThreadId = GetWindowThreadProcessId(hwnd, None)
        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)

def chat_last_save():
    getText = get_chat()
    getText = getText.split('\r\n')
    getText = pd.DataFrame(getText)
    getText[0] = getText[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')
    return getText.index[-2], getText.iloc[-2, 0]

def chat_check_command(cls, clst):
    getText = get_chat()
    getText = getText.split('\r\n')
    getText = pd.DataFrame(getText)
    getText[0] = getText[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')

    if getText.iloc[-2, 0] != clst:
        getText_ = getText.iloc[cls+1:, 0]
        getStr = str(getText_)

        check = 0
        s1 = str(); s2 = str(); s3 = str(); s4 = str(); s5 = str();
        r1 = 0; r2 = 0; r3 = 0; r4 = 0; r5 = 0;
        for line in lines:
            line = line.strip("\n")
            if check == 1:
                s1 = line
                r1 = SequenceMatcher(None, getStr, s1).ratio()
            if check == 2:
                s2 = line
                r2 = SequenceMatcher(None, getStr, s2).ratio()
            if check == 3:
                s3 = line
                r3 = SequenceMatcher(None, getStr, s3).ratio()
            if check == 4:
                s4 = line
                r4 = SequenceMatcher(None, getStr, s4).ratio()
            if check == 5:
                s5 = line
                r5 = SequenceMatcher(None, getStr, s5).ratio()
            if check >= 1:
                check += 1
            if check > 5:
                print(r1, r2, r3, r4, r5)
                printStr = str()
                if r1 >= r2 and r1 >= r3 and r1 >= r4 and r1 >= r5:
                    printStr = s1
                elif r2 >= r1 and r2 >= r3 and r2 >= r4 and r2 >= r5:
                    printStr = s2
                elif r3 >= r1 and r3 >= r2 and r3 >= r4 and r3 >= r5:
                    printStr = s3
                elif r4 >= r1 and r4 >= r2 and r4 >= r2 and r4 >= r5:
                    printStr = s4
                elif r5 >= r1 and r5 >= r2 and r5 >= r3 and r5 >= r4:
                    printStr = s5
                kakao_sendtext(printStr.strip("\n"))
                break
            val = getStr.find(line)
            if val != -1:
                check += 1

def main():
    print("Total chat data count : " + str(len(lines)))
    cls, clst = chat_last_save()

    while True:
        chat_check_command(cls, clst)
        cls, clst = chat_last_save()
        time.sleep(10)

if __name__ == '__main__':
    main()