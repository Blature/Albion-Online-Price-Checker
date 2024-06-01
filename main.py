import pyautogui
import pytesseract
import os
import pygetwindow as gw
import time
import sys
import gspread
from time import sleep
from oauth2client.service_account import ServiceAccountCredentials


os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "src/cred.json"

class WindowNotFoundException(Exception):
    pass

def win():
    window_title = 'Albion Online Client'


    window = gw.getWindowsWithTitle(window_title)

    if window:

        app_window = window[0]
        app_window.activate()
        time.sleep(1)  
    else:
        raise WindowNotFoundException(f"Window with Title '{window_title}' Not Found.")
    
try:
    win()
except WindowNotFoundException as e:
    print(e)
    sys.exit(1)

def find_tier(tier: int):
    x = 1000
    y = 0
    y_coordinates = [300, 325, 350, 375, 400, 425, 450, 475, 500]
    if 0 <= tier < len(y_coordinates):
        y = y_coordinates[tier]
    else:
        raise ValueError("Invalid tier value")
    
    return x, y
    
def find_ench(ench: int):
    x = 1100
    y = 0
    y_coordinates = [325, 350, 375, 400, 425, 450]
    if 0 <= ench < len(y_coordinates):
        y = y_coordinates[ench]
    else:
        raise ValueError("Invalid enchantment value")
    
    return x, y

def find_quality(quality: str):
    x = 1200
    y_coordinates = {
        "n": 325,
        "normal": 325,
        "g": 350,
        "good": 350,
        "o": 375,
        "outstanding": 375,
        "e": 400,
        "excellent": 400,
        "m": 425,
        "masterpiece": 425
    }
    
    y = y_coordinates.get(quality)
    if y is None:
        raise ValueError("Invalid quality value")
    
    return x, y

def search_item(name: str, tier: int, ench: int = None, quality: str = None):
    pyautogui.moveTo(705, 275, duration=0.5)
    pyautogui.click()
    pyautogui.moveTo(631, 275, duration=0.5)
    pyautogui.click()

    pyautogui.typewrite(name, interval=0.1)

    x,y = find_tier(tier)
    pyautogui.moveTo(1000, 275, duration=0.5)
    pyautogui.click()
    pyautogui.moveTo(x,y, duration=0.5)
    pyautogui.click()

    if ench is not None:
        xx,yy = find_ench(ench)
    else:
        xx,yy = 1100, 300
    pyautogui.moveTo(1100, 275, duration=0.5)
    pyautogui.click()
    pyautogui.moveTo(xx,yy, duration=0.5)
    pyautogui.click()

    if quality is not None:
        xxx,yyy = find_quality(quality)
        pyautogui.moveTo(1200, 275, duration=0.5)
        pyautogui.click()
        pyautogui.moveTo(xxx,yyy, duration=0.5)
        pyautogui.click()
    pyautogui.moveTo(705,275, duration=0.5)

def clean_text(text):
        text = text.strip()  # Trim whitespace
        text = text.replace('\\', '')  # Remove backslashes
        text = ' '.join(text.split())
        return text

def image(item_name):
    directory = f"src/{item_name}/"
    if not os.path.exists(directory):
        os.makedirs(directory)

    screenshot_path = f"{directory}/{item_name}.png"
    x, y, width, height = 658, 400, 530, 70
    screenshot = pyautogui.screenshot(region=(x, y, width, height))
    screenshot.save(screenshot_path)

    item_name_x, item_name_y, item_name_width, item_name_height = 0, 0, 240, 70
    item_name_area = screenshot.crop((item_name_x, item_name_y, item_name_x + item_name_width, item_name_y + item_name_height))

    item_value_x, item_value_y, item_value_width, item_value_height = 410, 0, 120, 70
    item_value_area = screenshot.crop((item_value_x, item_value_y, item_value_x + item_value_width, item_value_y + item_value_height))

    grayscale_image_name = item_name_area.convert('L')
    grayscale_image_value = item_value_area.convert('L')
    threshold = 128
    binary_image_name = grayscale_image_name.point(lambda p: p > threshold and 255)
    binary_image_value = grayscale_image_value.point(lambda p: p > threshold and 255)

    text_name = pytesseract.image_to_string(binary_image_name)
    text_value = pytesseract.image_to_string(binary_image_value)

    

    print(clean_text(text_name))
    print(clean_text(text_value))
    return {clean_text(text_name): clean_text(text_value)}


def update_google_sheet(sheet, worksheet, name, value, row, town_column, value_column):
    # Update the specified cell with the given name or value
    worksheet.update_cell(row, town_column, name)  # Item name column
    worksheet.update_cell(row, value_column, value)  # Specified town column

def get_base_row(item_name):

    item_base_rows = {
        'Cloth': 3,
        'Plank': 25,
        'Steel': 47,
        'Leather': 69
    }
    return item_base_rows.get(item_name, 3) 

def update_google_sheet_batch(sheet_name, worksheet_name, data_list, base_row, town_column, value_column):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('src/cred.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)

    # Iterate over the data list and update the Google Sheet
    for i, data in enumerate(data_list):
        row = base_row + i
        for name, value in data.items():
            update_google_sheet(sheet, worksheet, name, value, row, town_column, value_column)

def run(item_name: str, column: int):

    base_row = get_base_row(item_name)
    
    town_column = 2  # Column B for names
    value_column = (column - 1) * 1 + 4  # Start from column D for values
    
    tiers = range(4, 9)
    enchants = range(4)

    results = []

    for tier in tiers:
        for enchant in enchants:
            search_item(item_name, tier, enchant)
            image_result = image(f"{item_name}_{tier}_{enchant}")
            results.append(image_result)

    update_google_sheet_batch('Albion Online Bot', 'Sheet1', results, base_row, town_column, value_column)

def read_text_file(file_path):
    result = []
    with open(file_path, 'r') as file: 
        for line in file:
            result.append(line.strip())
    return result

def clean_fields():
    pyautogui.moveTo(705, 275, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(800, 275, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(800, 300, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(1000, 275, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(1000, 300, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(1100, 275, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(1100, 300, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(1200, 275, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(1200, 300, duration=0.1)
    pyautogui.click()
    pyautogui.moveTo(631, 275, duration=0.1)

def take_image():
    x, y, width, height = 658, 400, 530, 70
    screenshot = pyautogui.screenshot(region=(x, y, width, height))
    item_value_x, item_value_y, item_value_width, item_value_height = 410, 0, 120, 70
    item_value_area = screenshot.crop((item_value_x, item_value_y, item_value_x + item_value_width, item_value_y + item_value_height))

    grayscale_image_value = item_value_area.convert('L')
    threshold = 128
    binary_image_value = grayscale_image_value.point(lambda p: p > threshold and 255)

    text_value = pytesseract.image_to_string(binary_image_value)
    return clean_text(text_value)
def search_by_text(array):
    clean_fields()
    prices = []
    for item_name in array:
        pyautogui.click()
        pyautogui.moveTo(631, 275, duration=0.1)
        pyautogui.click()
        pyautogui.typewrite(item_name, interval=0.1)
        sleep(1.1)
        value = take_image()
        prices.append(value)
        pyautogui.moveTo(705, 275, duration=0.1)
    return prices

def cell_to_indices(cell):
    col = 0
    row = 0
    for char in cell:
        if char.isdigit():
            row = row * 10 + int(char)
        else:
            col = col * 26 + (ord(char.upper()) - ord('A') + 1)
    return row, col

def update_google_sheet(sheet_name, worksheet_name, start_cell, direction, prices):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('src/cred.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)

    row, col = cell_to_indices(start_cell)
    for price in prices:
        worksheet.update_cell(row, col, price)
        if direction == "Vertical":
            row += 1
        elif direction == "Horizontal":
            col += 1

def search_by_text_col_row(text_file, sheet, cell, direction):
    prices = search_by_text(read_text_file(text_file))
    print(prices)
    update_google_sheet('Albion Online Bot', sheet, cell, direction, prices)

search_by_text_col_row('src/Potion.txt', 'Sheet1', 'D92', 'Vertical')
# run('leather', 4) 

# search_item('cloth', 4, 1)
# image('test')
# run('Cloth', 1)

# 1.Martlock
# 2.Bridgewatch
# 3.Lymhurst
# 4.Fort Sterling
# 5.Thetford

#template for each item 