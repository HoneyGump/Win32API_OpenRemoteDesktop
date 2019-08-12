#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32api,win32gui,win32con
import win32clipboard as wc
import time

str_ip = ['10.0.0.17','10.0.0.19']
str_na = ['10.0.0.17 - 远程桌面连接','10.0.0.19 - 远程桌面连接']

name_user = win32api.GetUserName()
print(name_user)


# In[2]:


# Set the mouse 
def Click(position_x, position_y):
    win32api.SetCursorPos([position_x,position_y])
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, position_x, position_y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, position_x, position_y,0,0) 


# In[3]:


# set the keyboard
def KeyBoard(VK):
    win32api.keybd_event(VK,0,0,0)
    time.sleep(0.2)
    win32api.keybd_event(VK,0,2,0)


# In[4]:


def KeyStr(str):
    str = str.upper()
    for num in str:
        KeyBoard(ord(num))


# In[5]:


# Set the paste
def Paste(data_w):
    wc.OpenClipboard()
    wc.SetClipboardText(data_w,wc.CF_UNICODETEXT)
    #data_r = wc.GetClipboardData(wc.CF_UNICODETEXT)
    wc.CloseClipboard()
    win32api.keybd_event(win32con.VK_CONTROL,0,0,0)
    win32api.keybd_event(ord('V'),0,0,0)
    win32api.keybd_event(win32con.VK_CONTROL,0,2,0)# 2 is equal to the KEYEVENTF_KEYUP
    win32api.keybd_event(ord('V'),0,2,0)

# In[]


# define the function about the wait for the handle 
def findwindw( wClass, wName):
    time_home = win32api.GetTickCount()
    while True :
        hWnd = win32gui.FindWindow( wClass, wName)
        if hWnd != 0:
            break 
        time_end = win32api.GetTickCount()
        time_all = time_end - time_home
        if time_all > 10000:
            break
    print(hex(hWnd))
    return hWnd

def findwindwEx(wHandle, cHandle, wClass, wName):
    time_home = win32api.GetTickCount()
    while True :
        hWnd = win32gui.FindWindowEx( wHandle, cHandle, wClass, wName)
        if hWnd != 0:
            break 
        time_end = win32api.GetTickCount()
        time_all = time_end - time_home
        if time_all > 10000:
            break
    print(hex(hWnd))
    return hWnd



# In[6]:


def OpenRemote(num):
    # Control SocksCap64
    hWnd = findwindw('SocksCap64_WND_Class_150213', None)
    c_hWnd =findwindwEx(hWnd,0,'Static','About SocksCap64')

    # win32gui.PostMessage(c_hWnd, win32con.BM_CLICK)
    Rect = win32gui.GetWindowRect(c_hWnd)
    print(Rect)

    # set front
    win32gui.ShowWindow(hWnd,win32con.SW_SHOWNA)
    KeyBoard(win32con.VK_DOWN)
    win32gui.SetForegroundWindow(hWnd)
    # Open remote
    # win32gui.PostMessage(c_hWnd,win32con.BM_CLICK)
    # KeyBoard(win32con.VK_LEFT)
    # KeyBoard(win32con.VK_RETURN)
    Click(Rect[0]+20,Rect[1]+20)
    KeyBoard(13)
    time.sleep(5)

    # select ip open and input password
    # select
    hWnd = findwindw('#32770','远程桌面连接')
    c_hWnd = findwindwEx(hWnd, None, 'Button', '连接(&N)')
    com_hWnd = win32gui.FindWindowEx(hWnd,0,'ComboBoxEx32',None)
    com2_hWnd = win32gui.FindWindowEx(com_hWnd,0,'ComboBox',None)
    e_hWnd = win32gui.FindWindowEx(com2_hWnd,0,'Edit',None)
    # 设置文本框内容，等窗口处理完毕后返回true。中文需编码成gbk 
    # 参数：句柄；消息类型；参数WParam，无需使用； 参数IParam，要设置的内容，字符串
    win32api.SendMessage(e_hWnd, win32con.WM_SETTEXT, 0, str_ip[num])
    print('Selected index is ', str_ip[num])
    
    # open
    win32api.PostMessage(c_hWnd,win32con.BM_CLICK)
    hwd = findwindw('TscShellContainerClass', str_na[num]) 
    win32gui.ShowWindow(hwd,win32con.SW_SHOWMAXIMIZED)
    KeyBoard(win32con.VK_DOWN)
    win32gui.SetForegroundWindow(hwd)
    
    # set Xvnc
    y_xorg = 575
    Click(1075,y_xorg)
    Click(908,y_xorg+40)
    
    # set password
    Click(905,627)
    time.sleep(0.5)
    KeyStr('chencheng123123')
    KeyBoard(13) #Enter

    # find child handle
    # hwndChildList = []     
    # win32gui.EnumChildWindows(hWnd, lambda hwnd, param: param.append(hwnd),  hwndChildList)
    # print(hwndChildList)

    # select combox
#     CB_handle = com2_hWnd
#     PCB_handle = com_hWnd
#     if win32api.SendMessage(CB_handle, win32con.CB_SETCURSEL,num, 0) == 1:
#     # 下选框的父窗口命令231
#     # 参数：父窗口句柄； 命令； 参数：WParam：高位表示类型，低位表示内容；参数IParam，下选框句柄
#     # CBN_SELENDOK当用户选择了有效的列表项时发送，提示父窗体处理用户的选择。 LOWORD为组合框的ID. HIWORD为CBN_SELENDOK的值。
#         win32api.SendMessage(PCB_handle, win32con.WM_COMMAND, 0x90000, CB_handle) 
#     # CBN_SELCHANGE当用户更改了列表项的选择时发送，不论用户是通过鼠标选择或是通过方向键选择都会发送此通知。LOWORD为组合框的ID. HIWORD为CBN_SELCHANGE的值。
#         win32api.SendMessage(PCB_handle, win32con.WM_COMMAND, 0x10000, CB_handle)
    # Show which one has been chosed
#     gg = win32api.SendMessage(PCB_handle, win32con.CB_GETCURSEL, None , None)
#     print(gg)


# In[]:


# open Xshell
if win32gui.FindWindow('Xshell6::MainFrame_0',None) == 0:
    win32api.ShellExecute(0, 'open', r'C:\Program Files (x86)\NetSarang\Xshell 6\Xshell.exe', '', '', 1)
    # Control Xshell
    hWnd = findwindw(None, '会话')
    hWnd_1 = findwindwEx(hWnd, None, 'Button', '连接(&C)')
    win32gui.PostMessage(hWnd_1, win32con.BM_CLICK)
else:
    print('Xshell had been opened.')


# In[]:

num = 0
# open remote in socksCap64
if win32gui.FindWindow('TscShellContainerClass', str_na[num]) == 0:
    # open SocksCap64
    win32api.ShellExecute(0, 'open', r'C:\Program Files\SocksCap64-4.7\SocksCap64_RunAsAdmin.exe', '', '', 1)
    time.sleep(2)
    OpenRemote(num)
    print('Done')
else:
    print('Remote Desktop had been opened.')



#%%
