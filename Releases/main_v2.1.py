import win32gui, win32con, win32api, ctypes, time
import pandas as pd
from pywinauto import clipboard

room = win32gui.FindWindow(None, "바위게")
inBox = win32gui.FindWindowEx(room, None , "RICHEDIT50W" , None)  # 채팅창의 메세지 입력창

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

    if getText.iloc[-2, 0] == clst:
        return getText.index[-2], getText.iloc[-2, 0]
    else:
        getText_ = getText.iloc[cls+1 : , 0]

        fp = open("QnA.txt", "rt", encoding='UTF8')
        lines = fp.read().splitlines()
        data = []

        for line in lines:
            data.append(line)

        for i in range(0, len(data), 2):
            found = getText_[getText_.str.contains(data[i])]
            if int(found.count()) >= 1:
                print(data[i])
                print(lines[i+1])
                kakao_sendtext(lines[i+1])

                time.sleep(1)
                getText = get_chat()
                getText = getText.split('\r\n')
                getText = pd.DataFrame(getText)
                getText[0] = getText[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')
        fp.close()

        return getText.index[-2], getText.iloc[-2, 0]

def main():
    cls, clst = chat_last_save()

    while True:
        cls, clst = chat_check_command(cls, clst)
        time.sleep(1)

if __name__ == '__main__':
    main()
