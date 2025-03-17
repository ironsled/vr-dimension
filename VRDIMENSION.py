import cv2
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import threading
import pygetwindow as gw
import mss
import numpy as np
import os
import sys
import ctypes
import webbrowser


LICENSE_TEXT = """
VR Dimension Open Source License

Copyright (c) 2025 
X Technologies llc
Sarasota, FL
software@computernightmares.com

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, or distribute copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

1. The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

2. The Software is provided \"as is\", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the Software or the use or other dealings in the Software.

Supporting the Project

VR Dimension is a passion project, made open and free for the community. If you find it helpful and would like to support its continued development, tips are appreciated but never expected. Your encouragement, feedback, and shared experiences are the best support of all!

Thank you for being part of the VR Dimension journey.

Sixx
"""


selected_window = None
exit_flag = False
capture_flag = False

def minimize_console():
    if os.name == 'nt':
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)

def find_msf_window():
    for w in gw.getAllWindows():
        if "Flight Simulator" in w.title:
            return w.title
    return None

def process_video():
    global selected_window, exit_flag, capture_flag
    
    # Create window name for consistency
    window_name = "VR Dimension ~Sixx"
    
    with mss.mss() as sct:
        while not exit_flag:
            if selected_window and capture_flag:
                try:
                    window = gw.getWindowsWithTitle(selected_window)[0]
                    bbox = {'top': window.top, 'left': window.left, 'width': window.width, 'height': window.height}
                    sct_img = sct.grab(bbox)
                    frame = np.array(sct_img)
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)

                    height, width = frame.shape[:2]
                    right_eye = frame[:, width // 2:]
                    eye_height, eye_width = right_eye.shape[:2]

                    crop_x = int(eye_width * 0.09)
                    crop_y_top = int(eye_height * 0.12)
                    crop_y_bottom = int(eye_height * 0.08)
                    cropped_eye = right_eye[crop_y_top:eye_height - crop_y_bottom, crop_x:eye_width - crop_x]

                    # Resize to 720x720 resolution
                    resized_eye = cv2.resize(cropped_eye, (720, 720), interpolation=cv2.INTER_AREA)
                    
                    # Create and update the window
                    cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
                    cv2.resizeWindow(window_name, 720, 720)
                    cv2.imshow(window_name, resized_eye)
                    
                    # Add a small delay to allow the window to refresh
                    key = cv2.waitKey(1)
                    if key & 0xFF == ord('q'):
                        break
                except IndexError:
                    # If window not found, wait a bit
                    cv2.waitKey(100)
                    continue
                except Exception as e:
                    print(f"Error: {e}")
                    cv2.waitKey(100)
                    continue
            else:
                # If not capturing, wait a bit
                cv2.waitKey(100)
    
    # Clean up
    cv2.destroyAllWindows()

def select_window(event):
    global selected_window
    selected_window = window_var.get()

def start_capture():
    global capture_flag
    capture_flag = True
    status_label.config(text="Status: Capturing")

def stop_capture():
    global capture_flag
    capture_flag = False
    status_label.config(text="Status: Stopped")

def refresh_windows():
    window_list = [w.title for w in gw.getAllWindows() if w.title]
    window_menu['values'] = window_list
    msf_window = find_msf_window()
    if msf_window:
        window_var.set(msf_window)
        global selected_window
        selected_window = msf_window

def open_link_in_thread(url):
    threading.Thread(target=lambda: webbrowser.open_new_tab(url), daemon=True).start()

def open_about_popup():
    popup = tk.Toplevel(root)
    popup.title("ABOUT VR DIMENSION")
    popup.geometry("800x600")

    main_frame = ttk.Frame(popup)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ttk.Label(main_frame, text="SUPPORT THE CREATOR", font=("Arial", 12, "bold")).pack(pady=(5, 2))

    coffee_link = ttk.Label(main_frame, text="BUY ME A COFFEE", foreground="blue", cursor="hand2")
    coffee_link.pack(pady=2)
    coffee_link.bind("<Button-1>", lambda e: open_link_in_thread("https://streamelements.com/survivewithsixx/tip"))

    twitch_link = ttk.Label(main_frame, text="VISIT MY TWITCH", foreground="blue", cursor="hand2")
    twitch_link.pack(pady=2)
    twitch_link.bind("<Button-1>", lambda e: open_link_in_thread("https://www.twitch.tv/survivewithsixx"))

    discord_link = ttk.Label(main_frame, text="JOIN DISCORD", foreground="blue", cursor="hand2")
    discord_link.pack(pady=2)
    discord_link.bind("<Button-1>", lambda e: open_link_in_thread("https://discord.gg/WuHNTWsyHt"))
    ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=5)

    text_frame = ttk.Frame(main_frame)
    text_frame.pack(fill="both", expand=True)

    text_widget = tk.Text(text_frame, wrap="word", height=60, width=180)
    text_widget.insert("1.0", "WHAT IS VR DIMENSION?\n\nAN APPLICATION TO CAPTURE ONE EYE VR DISPLAY FOR VIRTUAL DESKTOP USERS WITHOUT THE NEED TO CAST FROM HEADSET TO STREAM SINGLE VIEW.\n\nWHY WAS THIS APPLICATION CREATED?\n\nI COULD NOT FIND A SOLUTION THAT WAS SATISFACTORY TO USE PREFERRED CODEC & OPENXR VDXR OPTIONS WITHOUT TASKING THE VR HEADSET.\nVR DIMENSION WAS A PROOF OF CONCEPT, FROM IDEA TO REALITY.\n\nLET THE BLUE SKY BE YOUR CANVAS, PAINT YOUR DREAMS. ~SIXX\n\n")
    text_widget.insert("end", LICENSE_TEXT)
    text_widget.insert("end", "\n\n\u2022 Special thanks to CowbellCody for BETA testing\n\u2022 Shout out to Tony & The SydSquadron Flight Sim Community")
    text_widget.config(state="disabled")

    scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    text_widget.pack(side="left", fill="both", expand=True)

    popup.update_idletasks()
    popup.geometry(f"{popup.winfo_reqwidth()}x{popup.winfo_reqheight()}")
    popup.minsize(800, 600)
    popup.maxsize(1200, 900)

def on_closing():
    global exit_flag
    exit_flag = True
    root.destroy()

def show_antivirus_warning():
    warning = tk.Toplevel(root)
    warning.title("Antivirus Information")
    warning.geometry("600x300")
    
    frame = ttk.Frame(warning, padding=15)
    frame.pack(fill="both", expand=True)
    
    ttk.Label(frame, text="⚠️ ANTIVIRUS INFORMATION ⚠️", 
              font=("Arial", 12, "bold")).pack(pady=(5, 15))
    
    message = (
        "VR Dimension may trigger false positive warnings due to screen capture technology.\n\n"
        "• PyInstaller packaging sometimes triggers false positives\n"
        "• No malicious code exists in this application\n\n"
    )
    
    text = tk.Text(frame, wrap="word", height=8, width=60)
    text.insert("1.0", message)
    text.config(state="disabled")
    
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    
    text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    ttk.Button(frame, text="Close", command=warning.destroy).pack(pady=15)

minimize_console()

root = tk.Tk()
root.title("VR Dimension ~Sixx")
root.geometry("800x600")
root.protocol("WM_DELETE_WINDOW", on_closing)  # Handle window close event
window_var = tk.StringVar()

# Create main frame
main_frame = ttk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Create control panel frame
control_frame = ttk.Frame(main_frame)
control_frame.pack(pady=10)

ttk.Label(control_frame, text="Select Process:").grid(row=0, column=0, padx=5, pady=5)
window_menu = ttk.Combobox(control_frame, textvariable=window_var, width=40)
window_menu.grid(row=0, column=1, padx=5, pady=5)
window_menu.bind('<<ComboboxSelected>>', select_window)

# Add buttons in a more organized way
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=5)

ttk.Button(button_frame, text="Refresh", command=refresh_windows).grid(row=0, column=0, padx=5, pady=5)
ttk.Button(button_frame, text="Start Capture", command=start_capture).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(button_frame, text="Stop Capture", command=stop_capture).grid(row=0, column=2, padx=5, pady=5)
ttk.Button(button_frame, text="Buy Sixx Coffee", command=open_about_popup).grid(row=0, column=3, padx=5, pady=5)
ttk.Button(button_frame, text="AV Info", command=show_antivirus_warning).grid(row=0, column=4, padx=5, pady=5)

# Status indicator
status_frame = ttk.Frame(main_frame)
status_frame.pack(pady=5)
status_label = ttk.Label(status_frame, text="Status: Ready")
status_label.pack()

# Instructions
instructions_frame = ttk.Frame(main_frame)
instructions_frame.pack(fill="both", expand=True, pady=10)

instructions = [
    "STEP 1: OPEN SIMULATOR IN WINDOW MODE",
    "STEP 2: CONNECT VR HEADSET",
    "STEP 3: START VR MODE IN SIMULATOR",
    "STEP 4: SELECT GAME IN PROCESS DROP DOWN, PRESS REFRESH BUTTON, PRESS START CAPTURE BUTTON",
    "CLICK X IN GUI TO EXIT APP",
    "TIP: RESIZE GAME WINDOW TO REMOVE LEFT AND RIGHT BORDERS FROM VIRTUAL DISPLAY"
]

for i, instruction in enumerate(instructions):
    ttk.Label(instructions_frame, text=instruction, wraplength=750, justify="center").pack(pady=5)


# Initialize window list
refresh_windows()

# Start video processing thread
video_thread = threading.Thread(target=process_video, daemon=True)
video_thread.start()

# Start the main loop
root.mainloop()
