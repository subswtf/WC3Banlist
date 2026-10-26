import time
import pytesseract
import cv2
import pyautogui
from grab_screen import grab_screen
import keyboard
import openpyxl
import os
import difflib
import ctypes
import sys
import csv
import io
import re
import urllib.request

# Make the script DPI-aware so screen coordinates match real pixels at any Windows display scale
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# Tesseract OCR path - Set your Path to your tesseract.exe here like in the example
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# Google Sheet with the ban list (columns: name, reason). Paste the sheet's normal link here.
# The sheet must be shared as "Anyone with the link - Viewer". Leave empty to use the Excel file only.
GOOGLE_SHEET_URL = ''

# Path to the Excel file with the ban list - used if GOOGLE_SHEET_URL is empty or the sheet can't be reached
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
banlist_path = os.path.join(SCRIPT_DIR, 'banlist.xlsx')

# ---- Settings -------------------------------------------------------------
# Screen regions (x1, y1, x2, y2) were made for a 3840x2160 (4K) screen. They are scaled
# automatically to your screen size. If names aren't found, press F4 and check the debug images.
BASE_WIDTH, BASE_HEIGHT = 3840, 2160
BASE_REGIONS = [(400, 520, 850, 900), (1700, 520, 2150, 900)]
# Click offsets (4K values, also scaled): click on the name, then on "Ban" in the menu below it
BASE_NAME_CLICK = (40, 15)
BASE_BAN_CLICK_Y = 260
# Fuzzy matching: only names at least this long are fuzzy-matched; shorter names must match exactly
FUZZY_MIN_LENGTH = 5
FUZZY_RATIO = 0.85
# Type the "Download Wc3Banlist..." advert into lobby chat on every F3
SHOW_ADVERT = False
# ---------------------------------------------------------------------------

SCREEN_W, SCREEN_H = pyautogui.size()
SX, SY = SCREEN_W / BASE_WIDTH, SCREEN_H / BASE_HEIGHT
REGIONS = [(int(a * SX), int(b * SY), int(c * SX), int(d * SY)) for a, b, c, d in BASE_REGIONS]

def add_row(banlist, row):
    if not row or row[0] is None:
        return
    name = str(row[0]).strip().lower()
    reason = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ''
    if name:
        banlist[name] = reason

# Load Banlist from the Google Sheet (as CSV)
def load_banlist_from_sheet():
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', GOOGLE_SHEET_URL)
    if not m:
        print("GOOGLE_SHEET_URL doesn't look like a Google Sheets link.")
        return None
    gid = re.search(r'[#&?]gid=(\d+)', GOOGLE_SHEET_URL)
    url = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={gid.group(1) if gid else '0'}"
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            text = resp.read().decode('utf-8-sig')
    except Exception as e:
        print(f"Couldn't load the Google Sheet ({e}). Is it shared as 'Anyone with the link'?")
        return None
    if text.lstrip().lower().startswith('<!doctype html') or text.lstrip().startswith('<html'):
        print("Google returned a sign-in page - share the sheet as 'Anyone with the link - Viewer'.")
        return None
    banlist = {}
    for row in list(csv.reader(io.StringIO(text)))[1:]:
        add_row(banlist, row)
    print(f"Loaded {len(banlist)} names from the Google Sheet.")
    return banlist

# Load Banlist from the Google Sheet, or the Excel file as a fallback
def load_banlist():
    if GOOGLE_SHEET_URL:
        banlist = load_banlist_from_sheet()
        if banlist is not None:
            return banlist
        print("Falling back to the Excel file.")
    if not os.path.exists(banlist_path):
        print("Ban list file not found.")
        return {}

    wb = openpyxl.load_workbook(banlist_path)
    ws = wb.active
    banlist = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        add_row(banlist, row)

    return banlist

def preprocess_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    img_resized = cv2.resize(gray, None, fx=1.0, fy=1.0)
    _, processed_img = cv2.threshold(img_resized, 150, 255, cv2.THRESH_BINARY)
    return processed_img

def get_text_data(img):
    preprocessed_img = preprocess_image(img)
    data = pytesseract.image_to_data(preprocessed_img, output_type=pytesseract.Output.DICT)
    return data

def check_banned(players, banlist):
    banned_players = []
    for player in players:
        player_lower = player['text'].lower()
        for banned_player in banlist:
            exact = player_lower == banned_player
            fuzzy = (len(player_lower) >= FUZZY_MIN_LENGTH and len(banned_player) >= FUZZY_MIN_LENGTH
                     and difflib.SequenceMatcher(None, player_lower, banned_player).ratio() >= FUZZY_RATIO)
            if exact or fuzzy:
                banned_players.append((player, banlist[banned_player]))
                break
    return banned_players

def extract_players_with_positions():
    players = []
    for coords in REGIONS:
        frame = grab_screen(coords)
        text_data = get_text_data(frame)
        for i in range(len(text_data['text'])):
            if text_data['text'][i].strip():
                players.append({
                    'text': text_data['text'][i],
                    'x': text_data['left'][i] + coords[0],
                    'y': text_data['top'][i] + coords[1]
                })
    return players

def click_banned_players(banned_players):
    for player, reason in banned_players:
        x, y = player['x'], player['y']
        x_right = x + int(BASE_NAME_CLICK[0] * SX)
        y_below = y + int(BASE_NAME_CLICK[1] * SY)
        pyautogui.moveTo(x_right, y_below)
        pyautogui.click()
        time.sleep(1)
        y_below_second_click = y + int(BASE_BAN_CLICK_Y * SY)
        pyautogui.moveTo(x_right, y_below_second_click)
        pyautogui.click()
        time.sleep(0.5)

def handle_exit():
    print("Exit program.")
    global running
    running = False

# Run banlist
def handle_f3():
    players_with_positions = extract_players_with_positions()

    # Output players
    print("Extracted Players:")
    for player in players_with_positions:
        print(player['text'])

    if SHOW_ADVERT:
        keyboard.write("Download Wc3Banlist at: https://github.com/RealFutureAI/WC3Banlist")
        time.sleep(.1)
        keyboard.send('enter')
        time.sleep(.1)
    keyboard.write("Players banned:")
    time.sleep(.1)
    keyboard.send('enter')
    time.sleep(.1)

    banned_from_list = check_banned(players_with_positions, load_banlist())
    for player, reason in banned_from_list:
        keyboard.write(f"{player['text']} - Reason: {reason}")
        time.sleep(.1)
        keyboard.send('enter')
        print(f"{player['text']} - Reason: {reason}")

    # Click on banned players
    click_banned_players(banned_from_list)


# Debug: save what the script sees, so you can check the regions line up with the names
def handle_f4():
    debug_dir = os.path.join(SCRIPT_DIR, 'debug')
    os.makedirs(debug_dir, exist_ok=True)
    full = cv2.cvtColor(grab_screen(), cv2.COLOR_RGB2BGR)
    for i, (a, b, c, d) in enumerate(REGIONS):
        cv2.rectangle(full, (a, b), (c, d), (0, 0, 255), 3)
        region = grab_screen((a, b, c, d))
        cv2.imwrite(os.path.join(debug_dir, f'region_{i + 1}_ocr_input.png'), preprocess_image(region))
    cv2.imwrite(os.path.join(debug_dir, 'screen_with_regions.png'), full)
    print(f"Screen {SCREEN_W}x{SCREEN_H}, regions {REGIONS}")
    print("Names read:", [p['text'] for p in extract_players_with_positions()])
    print(f"Debug images saved to {debug_dir}")

# Add Hotkeys
keyboard.add_hotkey("F3", handle_f3)
keyboard.add_hotkey("F4", handle_f4)
keyboard.add_hotkey("F6", handle_exit)

# Variable to check if the program is running
running = True

print(f"Screen {SCREEN_W}x{SCREEN_H}, banlist: {GOOGLE_SHEET_URL or banlist_path}")
print("The Program runs in background. Press F3 to check banned, F4 for debug images, F6 to stop the program.")

# check if hotkeys are pressed
while running:
    time.sleep(0.1)  # short delay, to reduce CPU lag
