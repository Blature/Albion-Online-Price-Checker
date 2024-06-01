

import pyautogui
import pygetwindow as gw
import time

def get_window_coordinates(window_title):
    window = gw.getWindowsWithTitle(window_title)
    if window:
        app_window = window[0]
        return app_window.left, app_window.top, app_window.width, app_window.height
    else:
        print(f"Window with title '{window_title}' not found.")
        return None

print("Move your mouse to the desired position. Press Ctrl-C to quit.")
window_title = "Albion Online Client"  # Replace with your application's window title

try:
    while True:
        window_coords = get_window_coordinates(window_title)
        if window_coords:
            win_x, win_y, win_width, win_height = window_coords

            x, y = pyautogui.position()

            relative_x = x - win_x
            relative_y = y - win_y
            position_str = f"Window relative X: {relative_x} Y: {relative_y}"
            print(position_str, end="\r", flush=True)   
            time.sleep(0.1)  
except KeyboardInterrupt:
    print("\nDone.")