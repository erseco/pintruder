
"""

pintruder v0.1.0
~~~~~~~~~~~~~~~~

2015 - Ernesto Serrano <info@ernesto.es>

Este programa se conecta a una hoja de Gooogle SpreadhSheet e introduce
las lineas en primavera

"""

from __future__ import print_function
#import httplib2
import os
import gspread
import sys

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import time
import subprocess
import ConfigParser
import ctypes

# Obtenemos el path del script
path = os.path.dirname(os.path.abspath(sys.argv[0])) + "\\"

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

try:
    config = ConfigParser.RawConfigParser()
    config.read(path + 'pintruder.ini')

    SHEET_KEY =  config.get('pintruder', 'SHEET_KEY')
    SHEET_GID =  config.get('pintruder', 'SHEET_GID')

    PRIMAVERA_PATH          = config.get('pintruder', 'PRIMAVERA_PATH')
    ERROR_WINDOW_TITLE      = config.get('pintruder', 'ERROR_WINDOW_TITLE')
    ERROR_WINDOW_SIZE_H     = config.getint('pintruder', 'ERROR_WINDOW_SIZE_H')
    ERROR_WINDOW_SIZE_W     = config.getint('pintruder', 'ERROR_WINDOW_SIZE_W')

    DATE_POSITION_DAY_X     = config.getint('pintruder', 'DATE_POSITION_DAY_X')
    DATE_POSITION_DAY_Y     = config.getint('pintruder', 'DATE_POSITION_DAY_Y')
    DATE_POSITION_MONTH_X   = config.getint('pintruder', 'DATE_POSITION_MONTH_X')
    DATE_POSITION_MONTH_Y   = config.getint('pintruder', 'DATE_POSITION_MONTH_Y')
    DATE_POSITION_YEAR_X    = config.getint('pintruder', 'DATE_POSITION_YEAR_X')
    DATE_POSITION_YEAR_Y    = config.getint('pintruder', 'DATE_POSITION_YEAR_Y')

    POSITION_FAVORITE_X     = config.getint('pintruder', 'POSITION_FAVORITE_X')
    POSITION_FAVORITE_Y     = config.getint('pintruder', 'POSITION_FAVORITE_Y')

    POSITION_FIRSTLINE_X    = config.getint('pintruder', 'POSITION_FIRSTLINE_X')
    POSITION_FIRSTLINE_Y    = config.getint('pintruder', 'POSITION_FIRSTLINE_Y')


except :
    print("Error al leer el archivo pintruder.ini")
    exit()




SCOPES = 'https://spreadsheets.google.com/feeds'
CLIENT_SECRET_FILE = path +'client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credential_path = path + 'pintruder.credentials.json'

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():

    credentials = get_credentials()

    gc = gspread.authorize(credentials)

    # Abrimos el documento (en la hoja deseada)
    wks = gc.open_by_key(SHEET_KEY).worksheet(SHEET_GID)


    # Obtenemos todos los registros de la hoja
    cells = wks.get_all_records()
    n_row = 1

    l = {}

    date = ""

    while True :

        if cells[n_row]["Fecha"] == "" :
            break

        if cells[n_row]["Introducido"] != "TRUE" :
            cells[n_row]["Row"] = n_row + 2 # Hay que sumarle 2 porque get_all_records() empieza desde la 2a fila

            if cells[n_row]["Fecha"] != date :
                date = cells[n_row]["Fecha"] # primera fecha
                l[date] = []

            l[date].append(cells[n_row])

        n_row +=1

    # Si tenemos lineas para introducir
    if len(l) > 0 :

        open_primavera()

        # Recorremos las fechas
        for date, rows in l.iteritems() :

            send_new_form()

            date_split = date.split("/")
            date_day   = date_split[0]
            date_month = date_split[1]
            date_year  = date_split[2]

            # Controlamos que los dias y meses tengan 2 cifras
            if int(date_day) < 10 : date_day = "0" + date_day
            if int(date_month) < 10 : date_month = "0" + date_month

            # Introducimos la fecha en primavera
            set_date(date_day, date_month, date_year)

            send_first_line()

            # Introducimos las lineas en primavera
            for row in rows :

                send_paste_tab(row["Articulo_ID"]) # article
                send_paste_tab(row["Proyecto_ID"]) # project
                send_paste_tab(str(row["Horas"])) # hours
                send_paste_tab(row["Observaciones"]) # comments

                if check_alert_window() :
                    print("Se mostro una advertencia en Primavera")
                    exit()


            send_save_form()

            # Marcamos como Introducidas las lineas
            for row in rows :
                n_row = row["Row"]
                wks.update_cell(n_row, 8, "TRUE") #La posicion 8 es la columan "Introducido"

        close_primavera()


def close_primavera() :
    print("cerrando primavera...")
    #call_macro('close_primavera.mcr')


def set_date(day, month, year) :
    send_date_day(day)
    send_date_month(month)
    send_date_year(year)


def open_primavera() :

    process = subprocess.Popen(PRIMAVERA_PATH)

    time.sleep(4)

    send_key(VK_RETURN)

    time.sleep(15)

    send_click(POSITION_FAVORITE_X, POSITION_FAVORITE_Y)

    time.sleep(3)

def send_new_form() :

    PressKey(VK_CONTROL)
    PressKey(KEY_N)
    ReleaseKey(KEY_N)
    ReleaseKey(VK_CONTROL)

    time.sleep(0.5)

    send_key(KEY_P)
    send_key(KEY_I)
    send_key(KEY_N)

    time.sleep(1)

    send_key(VK_RETURN)

def send_save_form() :

    PressKey(VK_CONTROL)
    PressKey(KEY_G)
    ReleaseKey(KEY_G)
    ReleaseKey(VK_CONTROL)

    time.sleep(1)


def send_date_day(day) :
    send_click(DATE_POSITION_DAY_X, DATE_POSITION_DAY_Y)
    time.sleep(0.5)
    for n in day :
        send_key(n_to_keycode[int(n)])
        time.sleep(0.5)
    time.sleep(1)

def send_date_month(month) :
    send_click(DATE_POSITION_MONTH_X, DATE_POSITION_MONTH_Y)
    time.sleep(0.5)
    for n in month :
        send_key(n_to_keycode[int(n)])
        time.sleep(0.5)
    time.sleep(1)

def send_date_year(year) :
    send_click(DATE_POSITION_YEAR_X, DATE_POSITION_YEAR_Y)
    time.sleep(0.5)
    for n in year :
        send_key(n_to_keycode[int(n)])
        time.sleep(0.5)
    time.sleep(1)

def send_first_line() :
    send_click(POSITION_FIRSTLINE_X, POSITION_FIRSTLINE_Y)
    time.sleep(1)

def send_paste_tab(value) :

    copy_to_clipboard(value)

    PressKey(VK_CONTROL)
    PressKey(KEY_V)
    ReleaseKey(KEY_V)
    ReleaseKey(VK_CONTROL)

    time.sleep(0.5)
    send_key(VK_TAB)
    time.sleep(1)


def send_key(key) :
    ReleaseKey(key)
    PressKey(key)

def send_click(x, y) :
    ctypes.windll.user32.SetCursorPos(x,y)
    ctypes.windll.user32.mouse_event(MOUSE_LEFTDOWN, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(MOUSE_LEFTUP, 0, 0, 0, 0)




EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))


class get_wnd_rect(ctypes.Structure):
    _fields_ = [('L', ctypes.c_int),
                ('T', ctypes.c_int),
                ('R', ctypes.c_int),
                ('B', ctypes.c_int)]


alert_window = False
titles = []
def foreach_window(hwnd, lParam):
    global alert_window
    if ctypes.windll.user32.IsWindowVisible(hwnd) and not alert_window :
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
        titles.append((hwnd, buff.value))

        rect = get_wnd_rect()

        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))

        crtW_h = rect.B - rect.T
        crtW_w = rect.R - rect.L

        if str(buff.value) == ERROR_WINDOW_TITLE and ERROR_WINDOW_SIZE_H == int(crtW_h) and ERROR_WINDOW_SIZE_W == int(crtW_w) :
            alert_window = True

    return True

def check_alert_window() :

    EnumWindows(EnumWindowsProc(foreach_window), 0)

    return alert_window


def copy_to_clipboard(text):
    GMEM_DDESHARE = 0x2000
    CF_UNICODETEXT = 13
    d = ctypes.windll # cdll expects 4 more bytes in user32.OpenClipboard(None)
    try:  # Python 2
        if not isinstance(text, unicode):
            text = text.decode('mbcs')
    except NameError:
        if not isinstance(text, str):
            text = text.decode('mbcs')
    d.user32.OpenClipboard(None)
    d.user32.EmptyClipboard()
    hCd = d.kernel32.GlobalAlloc(GMEM_DDESHARE, len(text.encode('utf-16-le')) + 2)
    pchData = d.kernel32.GlobalLock(hCd)
    ctypes.cdll.msvcrt.wcscpy(ctypes.c_wchar_p(pchData), text)
    d.kernel32.GlobalUnlock(hCd)
    d.user32.SetClipboardData(CF_UNICODETEXT, hCd)
    d.user32.CloseClipboard()


LONG = ctypes.c_long
DWORD = ctypes.c_ulong
ULONG_PTR = ctypes.POINTER(DWORD)
WORD = ctypes.c_ushort

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (('dx', LONG),
                ('dy', LONG),
                ('mouseData', DWORD),
                ('dwFlags', DWORD),
                ('time', DWORD),
                ('dwExtraInfo', ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (('wVk', WORD),
                ('wScan', WORD),
                ('dwFlags', DWORD),
                ('time', DWORD),
                ('dwExtraInfo', ULONG_PTR))

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (('uMsg', DWORD),
                ('wParamL', WORD),
                ('wParamH', WORD))

class _INPUTunion(ctypes.Union):
    _fields_ = (('mi', MOUSEINPUT),
                ('ki', KEYBDINPUT),
                ('hi', HARDWAREINPUT))

class INPUT(ctypes.Structure):
    _fields_ = (('type', DWORD),
                ('union', _INPUTunion))

def SendInput(*inputs):
    nInputs = len(inputs)
    LPINPUT = INPUT * nInputs
    pInputs = LPINPUT(*inputs)
    cbSize = ctypes.c_int(ctypes.sizeof(INPUT))
    return ctypes.windll.user32.SendInput(nInputs, pInputs, cbSize)

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARD = 2

def Input(structure):
    if isinstance(structure, MOUSEINPUT):
        return INPUT(INPUT_MOUSE, _INPUTunion(mi=structure))
    if isinstance(structure, KEYBDINPUT):
        return INPUT(INPUT_KEYBOARD, _INPUTunion(ki=structure))
    if isinstance(structure, HARDWAREINPUT):
        return INPUT(INPUT_HARDWARE, _INPUTunion(hi=structure))
    raise TypeError('Cannot create INPUT structure!')


MOUSE_LEFTDOWN = 0x0002     # left button down
MOUSE_LEFTUP = 0x0004       # left button up

VK_LBUTTON = 0x01               # Left mouse button
VK_RBUTTON = 0x02               # Right mouse button
VK_CANCEL = 0x03                # Control-break processing
VK_MBUTTON = 0x04               # Middle mouse button (three-button mouse)
VK_XBUTTON1 = 0x05              # X1 mouse button
VK_XBUTTON2 = 0x06              # X2 mouse button
VK_BACK = 0x08                  # BACKSPACE key
VK_TAB = 0x09                   # TAB key
VK_CLEAR = 0x0C                 # CLEAR key
VK_RETURN = 0x0D                # ENTER key
VK_SHIFT = 0x10                 # SHIFT key
VK_CONTROL = 0x11               # CTRL key
VK_MENU = 0x12                  # ALT key
VK_PAUSE = 0x13                 # PAUSE key
VK_CAPITAL = 0x14               # CAPS LOCK key
VK_KANA = 0x15                  # IME Kana mode
VK_HANGUL = 0x15                # IME Hangul mode
VK_JUNJA = 0x17                 # IME Junja mode
VK_FINAL = 0x18                 # IME final mode
VK_HANJA = 0x19                 # IME Hanja mode
VK_KANJI = 0x19                 # IME Kanji mode
VK_ESCAPE = 0x1B                # ESC key
VK_CONVERT = 0x1C               # IME convert
VK_NONCONVERT = 0x1D            # IME nonconvert
VK_ACCEPT = 0x1E                # IME accept
VK_MODECHANGE = 0x1F            # IME mode change request
VK_SPACE = 0x20                 # SPACEBAR
VK_PRIOR = 0x21                 # PAGE UP key
VK_NEXT = 0x22                  # PAGE DOWN key
VK_END = 0x23                   # END key
VK_HOME = 0x24                  # HOME key
VK_LEFT = 0x25                  # LEFT ARROW key
VK_UP = 0x26                    # UP ARROW key
VK_RIGHT = 0x27                 # RIGHT ARROW key
VK_DOWN = 0x28                  # DOWN ARROW key
VK_SELECT = 0x29                # SELECT key
VK_PRINT = 0x2A                 # PRINT key
VK_EXECUTE = 0x2B               # EXECUTE key
VK_SNAPSHOT = 0x2C              # PRINT SCREEN key
VK_INSERT = 0x2D                # INS key
VK_DELETE = 0x2E                # DEL key
VK_HELP = 0x2F                  # HELP key
VK_LWIN = 0x5B                  # Left Windows key (Natural keyboard)
VK_RWIN = 0x5C                  # Right Windows key (Natural keyboard)
VK_APPS = 0x5D                  # Applications key (Natural keyboard)
VK_SLEEP = 0x5F                 # Computer Sleep key
VK_NUMPAD0 = 0x60               # Numeric keypad 0 key
VK_NUMPAD1 = 0x61               # Numeric keypad 1 key
VK_NUMPAD2 = 0x62               # Numeric keypad 2 key
VK_NUMPAD3 = 0x63               # Numeric keypad 3 key
VK_NUMPAD4 = 0x64               # Numeric keypad 4 key
VK_NUMPAD5 = 0x65               # Numeric keypad 5 key
VK_NUMPAD6 = 0x66               # Numeric keypad 6 key
VK_NUMPAD7 = 0x67               # Numeric keypad 7 key
VK_NUMPAD8 = 0x68               # Numeric keypad 8 key
VK_NUMPAD9 = 0x69               # Numeric keypad 9 key
VK_MULTIPLY = 0x6A              # Multiply key
VK_ADD = 0x6B                   # Add key
VK_SEPARATOR = 0x6C             # Separator key
VK_SUBTRACT = 0x6D              # Subtract key
VK_DECIMAL = 0x6E               # Decimal key
VK_DIVIDE = 0x6F                # Divide key
VK_F1 = 0x70                    # F1 key
VK_F2 = 0x71                    # F2 key
VK_F3 = 0x72                    # F3 key
VK_F4 = 0x73                    # F4 key
VK_F5 = 0x74                    # F5 key
VK_F6 = 0x75                    # F6 key
VK_F7 = 0x76                    # F7 key
VK_F8 = 0x77                    # F8 key
VK_F9 = 0x78                    # F9 key
VK_F10 = 0x79                   # F10 key
VK_F11 = 0x7A                   # F11 key
VK_F12 = 0x7B                   # F12 key
VK_F13 = 0x7C                   # F13 key
VK_F14 = 0x7D                   # F14 key
VK_F15 = 0x7E                   # F15 key
VK_F16 = 0x7F                   # F16 key
VK_F17 = 0x80                   # F17 key
VK_F18 = 0x81                   # F18 key
VK_F19 = 0x82                   # F19 key
VK_F20 = 0x83                   # F20 key
VK_F21 = 0x84                   # F21 key
VK_F22 = 0x85                   # F22 key
VK_F23 = 0x86                   # F23 key
VK_F24 = 0x87                   # F24 key
VK_NUMLOCK = 0x90               # NUM LOCK key
VK_SCROLL = 0x91                # SCROLL LOCK key
VK_LSHIFT = 0xA0                # Left SHIFT key
VK_RSHIFT = 0xA1                # Right SHIFT key
VK_LCONTROL = 0xA2              # Left CONTROL key
VK_RCONTROL = 0xA3              # Right CONTROL key


KEY_0 = 0x30
KEY_1 = 0x31
KEY_2 = 0x32
KEY_3 = 0x33
KEY_4 = 0x34
KEY_5 = 0x35
KEY_6 = 0x36
KEY_7 = 0x37
KEY_8 = 0x38
KEY_9 = 0x39
KEY_A = 0x41
KEY_B = 0x42
KEY_C = 0x43
KEY_D = 0x44
KEY_E = 0x45
KEY_F = 0x46
KEY_G = 0x47
KEY_H = 0x48
KEY_I = 0x49
KEY_J = 0x4A
KEY_K = 0x4B
KEY_L = 0x4C
KEY_M = 0x4D
KEY_N = 0x4E
KEY_O = 0x4F
KEY_P = 0x50
KEY_Q = 0x51
KEY_R = 0x52
KEY_S = 0x53
KEY_T = 0x54
KEY_U = 0x55
KEY_V = 0x56
KEY_W = 0x57
KEY_X = 0x58
KEY_Y = 0x59
KEY_Z = 0x5A

n_to_keycode = {0 : KEY_0,
           1 : KEY_1,
           2 : KEY_2,
           3 : KEY_3,
           4 : KEY_4,
           5 : KEY_5,
           6 : KEY_6,
           7 : KEY_7,
           8 : KEY_8,
           9 : KEY_9,
}


def KeybdInput(code, flags):
    return KEYBDINPUT(code, code, flags, 0, None)

def HardwareInput(message, parameter):
    return HARDWAREINPUT(message & 0xFFFFFFFF,
                         parameter & 0xFFFF,
                         parameter >> 16 & 0xFFFF)

def Mouse(flags, x=0, y=0, data=0):
    return Input(MouseInput(flags, x, y, data))

def Keyboard(code, flags=0):
    return Input(KeybdInput(code, flags))

def Hardware(message, parameter=0):
    return Input(HardwareInput(message, parameter))



# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)
class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time",ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]


def PressKey(hexKeyCode):

    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( hexKeyCode, 0x48, 0, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):

    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( hexKeyCode, 0x48, 0x0002, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


if __name__ == '__main__':
    main()