import cv2
import tkinter as tk
from tkinter import ttk, Scale, HORIZONTAL, messagebox, Label
import threading
import pygetwindow as gw
import mss
import numpy as np
import os
import ctypes
import time
import json
import logging
import webbrowser

# Configuration file
CONFIG_FILE = "config.json"

# Default settings
DEFAULT_CONFIG = {
    "resolution": "320x240",
    "frame_rate": 60,
    "brightness": 1.0,
    "contrast": 1.0,
    "crop_x": 0.1,
    "crop_y_top": 0.15,
    "crop_y_bottom": 0.3,
}

# Global variables
selected_window = None
exit_flag = False
capture_flag = False
config = DEFAULT_CONFIG.copy()
window_title = "VR Dimension ~Sixx"
video_thread = None  # Declare video_thread globally
status_label = None

# Logging setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_config():
    """Loads configuration from a JSON file."""
    global config
    try:
        with open(CONFIG_FILE, "r") as f:
            config.update(json.load(f))
        logging.info("Configuration loaded successfully.")
    except FileNotFoundError:
        logging.warning("Config file not found. Using defaults.")
    except json.JSONDecodeError:
        logging.error("Invalid config file. Using defaults.")

def save_config():
    """Saves configuration to a JSON file."""
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=4)  # Add indent for better readability
        logging.info("Configuration saved successfully.")
    except Exception as e:
        logging.error(f"Failed to save config: {e}")

def minimize_console():
    """Minimizes the console window on Windows."""
    if os.name == 'nt':
        try:
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
            logging.info("Console window minimized.")
        except Exception as e:
            logging.error(f"Failed to minimize console: {e}")

def find_msf_window():
    """Finds the Microsoft Flight Simulator window."""
    for w in gw.getAllWindows():
        if "Flight Simulator" in w.title:
            logging.info(f"MSFS window found: {w.title}")
            return w.title
    logging.info("MSFS window not found.")
    return None

def capture_screen(sct, bbox):
    """Captures a screenshot of the specified region."""
    try:
        sct_img = sct.grab(bbox)
        frame = np.array(sct_img)
        return cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
    except mss.exception.ScreenShotError as e:
        logging.error(f"Screen capture failed: {e}")
        return None
    except Exception as e:
        logging.error(f"Error during screen capture: {e}")
        return None

def crop_and_resize(frame, capture_resolution, crop_params):
    """Crops and resizes the captured frame."""
    try:
        height, width = frame.shape[:2]
        right_eye = frame[:, width // 2:]
        eye_height, eye_width = right_eye.shape[:2]

        crop_x = int(eye_width * crop_params["crop_x"])
        crop_y_top = int(eye_height * crop_params["crop_y_top"])
        crop_y_bottom = int(eye_height * crop_params["crop_y_bottom"])
        cropped_eye = right_eye[crop_y_top:eye_height - crop_y_bottom, crop_x:eye_width - crop_x]

        target_aspect = capture_resolution[0] / capture_resolution[1]
        cropped_aspect = cropped_eye.shape[1] / cropped_eye.shape[0]

        if cropped_aspect > target_aspect:
            new_width = int(cropped_eye.shape[0] * target_aspect)
            offset = (cropped_eye.shape[1] - new_width) // 2
            cropped_eye = cropped_eye[:, offset:offset + new_width]
        else:
            new_height = int(cropped_eye.shape[1] / target_aspect)
            offset = (cropped_eye.shape[0] - new_height) // 2
            cropped_eye = cropped_eye[offset:offset + new_height, :]

        resized_eye = cv2.resize(cropped_eye, capture_resolution, interpolation=cv2.INTER_LINEAR)
        return resized_eye
    except Exception as e:
        logging.error(f"Error during crop and resize: {e}")
        return None

def apply_adjustments(frame, brightness, contrast):
    """Applies brightness and contrast adjustments to the frame."""
    try:
        return cv2.convertScaleAbs(frame, alpha=contrast, beta=(brightness - 1.0) * 255)
    except Exception as e:
        logging.error(f"Error applying adjustments: {e}")
        return None

def display_frame(frame, window_name, capture_resolution):
    """Displays the processed frame."""
    try:
        cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
        cv2.resizeWindow(window_name, *capture_resolution)
        cv2.imshow(window_name, frame)
    except cv2.error as e:
        logging.error(f"OpenCV error displaying frame: {e}")
    except Exception as e:
        logging.error(f"Error displaying frame: {e}")

def process_video():
    """Main function for video processing."""
    global selected_window, exit_flag, capture_flag, config
    last_time = time.time()
    window_name = window_title
    with mss.mss() as sct:
        while not exit_flag:
            if selected_window and capture_flag:
                try:
                    windows = gw.getWindowsWithTitle(selected_window)
                    if not windows:
                        logging.warning(f"Window '{selected_window}' not found. Stopping capture.")
                        capture_flag = False
                        status_label.config(text="Status: Window Not Found")
                        continue

                    window = windows[0]
                    bbox = {'top': window.top, 'left': window.left, 'width': window.width, 'height': window.height}

                    frame = capture_screen(sct, bbox)
                    if frame is not None:
                        resolution = tuple(map(int, config["resolution"].split("x")))
                        cropped_eye = crop_and_resize(frame, resolution, config)
                        if cropped_eye is not None:
                            adjusted_frame = apply_adjustments(cropped_eye, config["brightness"], config["contrast"])
                            if adjusted_frame is not None:
                                display_frame(adjusted_frame, window_name, resolution)

                    if cv2.waitKey(1) & 0xFF == ord('q'):
                        break

                    # Control frame rate
                    current_time = time.time()
                    elapsed_time = current_time - last_time
                    target_delay = 1.0 / config["frame_rate"]
                    if elapsed_time < target_delay:
                        time.sleep(target_delay - elapsed_time)
                    last_time = current_time

                except Exception as e:
                    logging.error(f"Exception in video processing loop: {e}")
                    time.sleep(1)
            else:
                cv2.waitKey(100)
    cv2.destroyAllWindows()
    logging.info("Video processing thread stopped.")

def select_window_callback(event):
    """Callback for window selection from Combobox."""
    global selected_window
    selected_window = window_var.get()
    logging.info(f"Selected window: {selected_window}")

def start_capture_callback():
    """Callback for Start Capture button."""
    global capture_flag, video_thread
    if not selected_window:
        logging.warning("No window selected. Cannot start capture.")
        status_label.config(text="Status: No Window Selected")
        return
    capture_flag = True
    status_label.config(text="Status: Capturing")
    logging.info("Capture started.")
    if video_thread is None or not video_thread.is_alive():
        video_thread = threading.Thread(target=process_video, daemon=True)
        video_thread.start()

def stop_capture_callback():
    """Callback for Stop Capture button."""
    global capture_flag
    capture_flag = False
    status_label.config(text="Status: Stopped")
    logging.info("Capture stopped.")

def refresh_windows_callback():
    """Callback for Refresh Windows button."""
    global selected_window
    window_list = [w.title for w in gw.getAllWindows() if w.title]
    window_menu['values'] = window_list
    msf_window = find_msf_window()
    if msf_window:
        window_var.set(msf_window)
        selected_window = msf_window  # Set selected_window here
    logging.info("Window list refreshed.")

def update_resolution_callback(event):
    """Callback for resolution selection."""
    config["resolution"] = resolution_var.get()
    save_config()
    logging.info(f"Resolution updated to: {config['resolution']}")

def update_frame_rate_callback(val):
    """Callback for frame rate slider."""
    config["frame_rate"] = int(val)
    save_config()
    logging.info(f"Frame rate updated to: {config['frame_rate']}")

def update_brightness_callback(val):
    """Callback for brightness slider."""
    config["brightness"] = float(val)
    save_config()
    logging.info(f"Brightness updated to: {config['brightness']}")

def update_contrast_callback(val):
    """Callback for contrast slider."""
    config["contrast"] = float(val)
    save_config()
    logging.info(f"Contrast updated to: {config['contrast']}")

def on_closing_callback():
    """Callback for window closing event."""
    global exit_flag, video_thread
    exit_flag = True
    save_config()
    logging.info("Application closing...")
    if video_thread and video_thread.is_alive():
        video_thread.join()  # Wait for the video thread to finish
    root.destroy()

def open_about_popup():
    """Opens an About popup window."""
    about_window = tk.Toplevel(root)
    about_window.title("About VR Dimension ~Sixx")

    about_label = Label(about_window, text="VR Dimension Version 1.3\n\nDeveloped by: [Sixx/X Technologies llc]\n", justify=tk.LEFT)
    about_label.pack()

    discord_label = Label(about_window, text="Discord: ", justify=tk.LEFT)
    discord_label.pack()
    discord_link = Label(about_window, text="https://discord.gg/WuHNTWsyHt", fg="blue", cursor="hand2", justify=tk.LEFT)
    discord_link.pack()
    discord_link.bind("<Button-1>", lambda e: webbrowser.open("https://discord.gg/WuHNTWsyHt"))

    twitch_label = Label(about_window, text="Twitch: ", justify=tk.LEFT)
    twitch_label.pack()
    twitch_link = Label(about_window, text="https://www.twitch.tv/survivewithsixx", fg="blue", cursor="hand2", justify=tk.LEFT)
    twitch_link.pack()
    twitch_link.bind("<Button-1>", lambda e: webbrowser.open("https://www.twitch.tv/survivewithsixx"))

    thanks_label = Label(about_window, text="\nSpecial thanks to CowbellCody for BETA testing.\n\nShout out to Syd Tony & The SydSquadron Flight Sim Community", justify=tk.CENTER)
    thanks_label.pack()

def show_antivirus_warning():
    """Shows an Antivirus Warning popup window."""
    messagebox.showwarning("Antivirus Warning",
                        "Some antivirus software may flag this application due to its screen capture functionality.  This is a false positive.  Please add this application to your antivirus whitelist if necessary.")

def open_coffee_link():
    """Opens the Buy Sixx Coffee link in the default web browser."""
    webbrowser.open("https://streamelements.com/survivewithsixx/tip")
    logging.info("Opening Buy Sixx Coffee link.")

minimize_console()
load_config()

root = tk.Tk()
root.title(window_title)
root.geometry("800x600")
root.protocol("WM_DELETE_WINDOW", on_closing_callback)

window_var = tk.StringVar()
resolution_var = tk.StringVar(value=config["resolution"])

main_frame = ttk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

control_frame = ttk.Frame(main_frame)
control_frame.pack(pady=10)

# Window selection
ttk.Label(control_frame, text="Select Window:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
window_menu = ttk.Combobox(control_frame, textvariable=window_var, width=40)
window_menu.grid(row=0, column=1, padx=5, pady=5)
window_menu.bind('<<ComboboxSelected>>', select_window_callback)

# Resolution selection
ttk.Label(control_frame, text="Resolution:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
resolution_menu = ttk.Combobox(control_frame, textvariable=resolution_var, values=["320x240", "480x320", "640x480", "720x720", "1280x720"], width=15)
resolution_menu.grid(row=1, column=1, padx=5, pady=5)
resolution_menu.bind('<<ComboboxSelected>>', update_resolution_callback)

# Frame rate slider
ttk.Label(control_frame, text="Frame Rate:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
frame_slider = Scale(control_frame, from_=5, to=120, orient=HORIZONTAL, command=update_frame_rate_callback)
frame_slider.set(config["frame_rate"])
frame_slider.grid(row=2, column=1, padx=5, pady=5)

# Brightness slider
ttk.Label(control_frame, text="Brightness:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
brightness_slider = Scale(control_frame, from_=0.0, to=2.0, resolution=0.1, orient=HORIZONTAL, command=update_brightness_callback)
brightness_slider.set(config["brightness"])
brightness_slider.grid(row=3, column=1, padx=5, pady=5)

# Contrast slider
ttk.Label(control_frame, text="Contrast:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
contrast_slider = Scale(control_frame, from_=0.0, to=2.0, resolution=0.1, orient=HORIZONTAL, command=update_contrast_callback)
contrast_slider.set(config["contrast"])
contrast_slider.grid(row=4, column=1, padx=5, pady=5)

# Add buttons in a more organized way
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=5)

ttk.Button(button_frame, text="Refresh", command=refresh_windows_callback).grid(row=0, column=0, padx=5, pady=5)
ttk.Button(button_frame, text="Start Capture", command=start_capture_callback).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(button_frame, text="Stop Capture", command=stop_capture_callback).grid(row=0, column=2, padx=5, pady=5)
ttk.Button(button_frame, text="Buy Sixx Coffee", command=open_coffee_link).grid(row=0, column=3, padx=5, pady=5) # Moved
ttk.Button(button_frame, text="AV Info", command=show_antivirus_warning).grid(row=0, column=4, padx=5, pady=5)
ttk.Button(button_frame, text="About", command=open_about_popup).grid(row=0, column=5, padx=5, pady=5) # Moved

# Status indicator
status_frame = ttk.Frame(main_frame)
status_frame.pack(pady=5)
status_label = ttk.Label(status_frame, text="Status: Ready")
status_label.pack()

# Instructions
instructions_frame = ttk.Frame(main_frame)
instructions_frame.pack(fill="both", expand=True, pady=10)

ttk.Label(instructions_frame, text="INSTRUCTIONS", font=("Arial", 12, "bold")).pack(pady=(10, 5))

instructions = [
    "STEP 1: OPEN SIMULATOR IN WINDOW MODE",
    "STEP 2: CONNECT VR HEADSET",
    "STEP 3: START VR MODE IN SIMULATOR",
    "STEP 4: SELECT GAME IN PROCESS DROP DOWN, PRESS REFRESH BUTTON, PRESS START CAPTURE BUTTON",
    "CLICK X IN GUI TO EXIT APP"

]

for i, instruction in enumerate(instructions):
    ttk.Label(instructions_frame, text=instruction, wraplength=750, justify="center").pack(pady=5)

# Initialize window list
refresh_windows_callback()

video_thread = threading.Thread(target=process_video, daemon=True) #moved here
video_thread.start() #moved here

root.mainloop()
