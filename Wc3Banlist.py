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
# Tuned for a 3v3 lobby (3 rows per team), measured at 2560x1440 and stored in 4K units
BASE_REGIONS = [(387, 525, 849, 750), (1683, 525, 2150, 750)]
# Click offsets (4K values, also scaled): click on the name, then on "Ban" in the menu below it
BASE_NAME_CLICK = (40, 15)
BASE_BAN_CLICK_Y = 260
# Fuzzy matching: only names at least this long are fuzzy-matched; shorter names must match exactly
FUZZY_MIN_LENGTH = 5
FUZZY_RATIO = 0.85
# OCR tuning: enlarge the name areas before reading (bigger = clearer, slightly slower),
# brightness cut-off for text (0-255), and dark-text-on-white (what Tesseract reads best)
OCR_SCALE = 3
OCR_THRESHOLD = 150
OCR_INVERT = True
# Tesseract page mode: 6 = a block of text lines (player list)
OCR_CONFIG = '--psm 6'
# Kick banned players automatically by clicking their slot (off: chat notification only)
AUTO_KICK = False
# Move this script's window onto your second monitor at start-up, so it doesn't cover the lobby
CONSOLE_ON_SECOND_MONITOR = True
# Type the "Download Wc3Banlist..." advert into lobby chat on every F3
SHOW_ADVERT = False
# ---------------------------------------------------------------------------

SCREEN_W, SCREEN_H = pyautogui.size()
SX, SY = SCREEN_W / BASE_WIDTH, SCREEN_H / BASE_HEIGHT
REGIONS = [(int(a * SX), int(b * SY), int(c * SX), int(d * SY)) for a, b, c, d in BASE_REGIONS]

TICKED = {'true', 'yes', 'y', 'x', '1', '\u2713', '\u2714'}

# Turn sheet rows (first row = headers) into {name: (reason, 'ban' or 'warn')}
# Columns are found by header name, so their order doesn't matter:
#   Player (or Name), Reason, Banned, Warning - other columns (Date, Admin...) are ignored.
# If there are no Banned/Warning columns, every listed player is banned.
def parse_rows(rows):
    rows = list(rows)
    if not rows:
        return {}
    headers = [str(h).strip().lower() if h is not None else '' for h in rows[0]]
    def col(*names):
        for n in names:
            if n in headers:
                return headers.index(n)
        return None
    name_c = col('player', 'name')
    name_c = 0 if name_c is None else name_c
    reason_c, ban_c, warn_c = col('reason'), col('banned', 'ban'), col('warning', 'warn')
    if reason_c is None and ban_c is None and warn_c is None and len(headers) > 1:
        reason_c = 1
    def cell(row, c):
        return str(row[c]).strip() if c is not None and c < len(row) and row[c] is not None else ''
    banlist = {}
    for row in rows[1:]:
        if not row:
            continue
        name = cell(row, name_c).lower()
        if not name:
            continue
        if ban_c is None and warn_c is None:
            status = 'ban'
        elif cell(row, ban_c).lower() in TICKED:
            status = 'ban'
        elif cell(row, warn_c).lower() in TICKED:
            status = 'warn'
        else:
            continue
        # a ban always wins over a warning if the same name is listed twice
        if banlist.get(name, ('', ''))[1] != 'ban':
            banlist[name] = (cell(row, reason_c), status)
    return banlist

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
    banlist = parse_rows(csv.reader(io.StringIO(text)))
    bans = sum(1 for r, st in banlist.values() if st == 'ban')
    print(f"Loaded {bans} bans and {len(banlist) - bans} warnings from the Google Sheet.")
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
    return parse_rows(ws.iter_rows(values_only=True))

def preprocess_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    img_resized = cv2.resize(gray, None, fx=OCR_SCALE, fy=OCR_SCALE, interpolation=cv2.INTER_CUBIC)
    mode = cv2.THRESH_BINARY_INV if OCR_INVERT else cv2.THRESH_BINARY
    _, processed_img = cv2.threshold(img_resized, OCR_THRESHOLD, 255, mode)
    # white border helps Tesseract with text near the edges
    processed_img = cv2.copyMakeBorder(processed_img, 10, 10, 10, 10, cv2.BORDER_CONSTANT,
                                       value=255 if OCR_INVERT else 0)
    return processed_img

def get_text_data(img):
    preprocessed_img = preprocess_image(img)
    data = pytesseract.image_to_data(preprocessed_img, config=OCR_CONFIG, output_type=pytesseract.Output.DICT)
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
                reason, status = banlist[banned_player]
                banned_players.append((player, reason, status))
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
                    # convert from the enlarged image back to real screen pixels
                    'x': int((text_data['left'][i] - 10) / OCR_SCALE) + coords[0],
                    'y': int((text_data['top'][i] - 10) / OCR_SCALE) + coords[1]
                })
    return players

def click_banned_players(banned_players):
    for player, reason, status in banned_players:
        if status != 'ban':
            continue
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
    matches = check_banned(players_with_positions, load_banlist())
    bans = [m for m in matches if m[2] == 'ban']
    warnings = [m for m in matches if m[2] == 'warn']

    def say(line):
        keyboard.write(line)
        time.sleep(.1)
        keyboard.send('enter')
        time.sleep(.1)
        print(line)

    if bans:
        say("Players banned:")
        for player, reason, status in bans:
            say(f"{player['text']} - Reason: {reason}" if reason else player['text'])
    # Warnings (yellow card): announced in lobby chat, but the player is not kicked
    for player, reason, status in warnings:
        say(f"{player['text']} - WARNING: {reason}" if reason else f"{player['text']} - WARNING")
    if not matches:
        print("No banned or warned players found.")

    # Click on banned players (only if AUTO_KICK is on)
    if AUTO_KICK:
        click_banned_players(bans)


# Debug: save what the script sees, so you can check the regions line up with the names
def handle_f4():
    debug_dir = os.path.join(SCRIPT_DIR, 'debug')
    os.makedirs(debug_dir, exist_ok=True)
    # main monitor only, so the boxes line up on multi-monitor setups
    full = cv2.cvtColor(grab_screen((0, 0, SCREEN_W - 1, SCREEN_H - 1)), cv2.COLOR_RGB2BGR)
    for i, (a, b, c, d) in enumerate(REGIONS):
        cv2.rectangle(full, (a, b), (c, d), (0, 0, 255), 3)
        region = grab_screen((a, b, c, d))
        cv2.imwrite(os.path.join(debug_dir, f'region_{i + 1}_ocr_input.png'), preprocess_image(region))
    cv2.imwrite(os.path.join(debug_dir, 'screen_with_regions.png'), full)
    print(f"Screen {SCREEN_W}x{SCREEN_H}, regions {REGIONS}")
    print("Names read:", [p['text'] for p in extract_players_with_positions()])
    print(f"Debug images saved to {debug_dir}")

# Move the command prompt / terminal window to the first non-main monitor
def move_console_to_second_monitor():
    try:
        import win32api, win32gui, win32con
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if not hwnd:
            return
        # In Windows Terminal the console window is hidden; its owner is the real terminal window
        owner = ctypes.windll.user32.GetAncestor(hwnd, 3)  # GA_ROOTOWNER
        if owner:
            hwnd = owner
        for monitor, _, _ in win32api.EnumDisplayMonitors():
            info = win32api.GetMonitorInfo(monitor)
            if info['Flags'] & 1:  # main monitor - skip
                continue
            left, top, right, bottom = info['Work']
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.MoveWindow(hwnd, left + 40, top + 40,
                                min(1200, right - left - 80), min(700, bottom - top - 80), True)
            return
        print("Only one monitor found - window not moved.")
    except Exception as e:
        print(f"Couldn't move the window to the second monitor ({e}).")

if CONSOLE_ON_SECOND_MONITOR:
    move_console_to_second_monitor()

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
