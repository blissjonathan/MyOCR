#-- include('examples/showgrabfullscreen.py') --#
import pyscreenshot as ImageGrab
from time import sleep
import csv
import cv2
try:
    import Image
except ImportError:
    import PIL.Image
import pytesseract
from googletrans import Translator
import turtle
from win32api import GetSystemMetrics
import pywintypes
import win32api, win32con, win32gui, win32ui
from tkinter import *
import tkinter as tk
import random
from threading import Thread
import image_slicer

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
# pytesseract.pytesseract.tesseract_cmd = r"F:\Tesseract-OCR\tesseract.exe"
curtext = ""
tlist = {}
windows = []
ws = {}
cts = {}
pts = {}
rt = 0
def main(trans,lf,tp,w,h,i,name):
    hInstance = win32api.GetModuleHandle()
    className = 'MyTranslations||' + str(i) + '||' + str(rt) + '||' + name

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633576(v=vs.85).aspx
    # win32gui does not support WNDCLASSEX.
    wndClass                = win32gui.WNDCLASS()
    # http://msdn.microsoft.com/en-us/library/windows/desktop/ff729176(v=vs.85).aspx
    wndClass.style          = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    wndClass.lpfnWndProc    = wndProc
    wndClass.hInstance      = hInstance
    wndClass.hCursor        = win32gui.LoadCursor(None, win32con.IDC_ARROW)
    wndClass.hbrBackground  = win32gui.GetStockObject(win32con.WHITE_BRUSH)
    wndClass.lpszClassName  = className
    # win32gui does not support RegisterClassEx
    wndClassAtom = win32gui.RegisterClass(wndClass)

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ff700543(v=vs.85).aspx
    # Consider using: WS_EX_COMPOSITED, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT
    # The WS_EX_TRANSPARENT flag makes events (like mouse clicks) fall through the window.
    exStyle = win32con.WS_EX_COMPOSITED | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ms632600(v=vs.85).aspx
    # Consider using: WS_DISABLED, WS_POPUP, WS_VISIBLE
    style = win32con.WS_DISABLED | win32con.WS_POPUP | win32con.WS_VISIBLE

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ms632680(v=vs.85).aspx
    hWindow = win32gui.CreateWindowEx(
        exStyle,
        wndClassAtom,
        'txt_' + str(i) + '_' + str(rt) + '_' + name, # WindowName
        style,
        lf, # x
        tp, # y
        w+20, # width
        h+20, # height
        None, # hWndParent
        None, # hMenu
        hInstance,
        None # lpParam
    )
    ws[name].append(hWindow)

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633540(v=vs.85).aspx
    win32gui.SetLayeredWindowAttributes(hWindow, 0x00FFFFFF, 255, win32con.LWA_COLORKEY | win32con.LWA_ALPHA)

    # http://msdn.microsoft.com/en-us/library/windows/desktop/dd145167(v=vs.85).aspx
    #win32gui.UpdateWindow(hWindow)

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633545(v=vs.85).aspx
    win32gui.SetWindowPos(hWindow, win32con.HWND_TOPMOST, 0, 0, 0, 0,
        win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)

    # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx
    #win32gui.ShowWindow(hWindow, win32con.SW_SHOW)

    win32gui.PumpMessages()

def wndProc(hWnd, message, wParam, lParam):
    if message == win32con.WM_PAINT:
        hdc, paintStruct = win32gui.BeginPaint(hWnd)

        dpiScale = win32ui.GetDeviceCaps(hdc, win32con.LOGPIXELSX) / 60.0
        fontSize = 18

        # http://msdn.microsoft.com/en-us/library/windows/desktop/dd145037(v=vs.85).aspx
        lf = win32gui.LOGFONT()
        lf.lfFaceName = "Times New Roman"
        lf.lfHeight = int(round(dpiScale * fontSize))
        #lf.lfWeight = 150
        # Use nonantialiased to remove the white edges around the text.
        # lf.lfQuality = win32con.NONANTIALIASED_QUALITY
        hf = win32gui.CreateFontIndirect(lf)
        win32gui.SelectObject(hdc, hf)

        rect = win32gui.GetClientRect(hWnd)

        cn = win32gui.GetClassName(hWnd)
        cnName = cn.split('||')[3]
        cnt = cts[cnName]

        # http://msdn.microsoft.com/en-us/library/windows/desktop/dd162498(v=vs.85).aspx
        win32gui.DrawText(
            hdc,
            cnt,
            -1,
            rect,
            win32con.DT_LEFT | win32con.DT_NOCLIP | win32con.DT_SINGLELINE | win32con.DT_TOP
        )
        win32gui.SetTextColor(hdc,win32api.RGB(0,0,0));
        win32gui.EndPaint(hWnd, paintStruct)
        return 0

    elif message == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0

    else:
        return win32gui.DefWindowProc(hWnd, message, wParam, lParam)

def processSlice(name,count):
    print("Starting Slice_" + name)
    if name in ws:
        for x in ws[name]:
            win32gui.PostMessage(x,win32con.WM_CLOSE,0,0)
    iwidth, iheight = PIL.Image.open(name).size
    ws[name] = []
    imdata = pytesseract.image_to_data(PIL.Image.open(name), lang="chi_tra", config='--disable-openmp --disable-shared', nice=0, output_type="dict")
    tops = 0;
    lefts = 0;
    widths = 0;
    heights = 0;
    confs = 0;
    lb = 0
    aw = 0
    ah = 0
    if count == 1:
        aw = iwidth
    if count == 2:
        ah = iheight
    if count == 3:
        aw = iwidth
        ah = iheight

    print(ah,aw,name)
    for key, value in imdata.items():
        if(key== "conf"):
            confs = value
        if(key == "height"):
            heights = value
        if(key == "width"):
            widths = value
        if(key == "top"):
            tops = value
        if(key == "left"):
            lefts = value
        if(key == "text"):
            i = 0
            print("translating..." + name)
            while i < len(value):
                txt = value[i]
                tp = tops[i] + ah
                lf = lefts[i] + aw
                w = widths[i]
                h = heights[i]
                if(confs[i] > 85):
                    if txt:
                        trans = translator.translate(txt, dest='en', src='zh-TW').text
                        cts[name] = trans
                        pt = Thread(target=main, args=[trans,lf,tp,w,h,i,name])
                        pt.start()
                        while not pt.is_alive():
                            pass
                i += 1
    sleep(5);


if __name__ == "__main__":
    while True:
        rt += 1
        translator = Translator()
        im=ImageGrab.grab()
        im.convert('L')
        im.save('sc.png')
        col = PIL.Image.open("sc.png")
        gray = col.convert('L')
        bw = gray.point(lambda x: 0 if x<128 else 255, '1')
        bw.save("sc2.png")
        slices = image_slicer.slice('sc2.png', 4)
        scount = 0
        for slice in slices:
            name = str(slice).split('- ')[1]
            name = name.replace(">", "")
            if (name) in pts:
                t = pts[name]
                if not t.is_alive():
                    pts[name] = Thread(target=processSlice, args=[name,scount])
                    pts[name].start()
            else:
                pts[name] = Thread(target=processSlice, args=[name,scount])
                pts[name].start()
            scount += 1

    sleep(5)
