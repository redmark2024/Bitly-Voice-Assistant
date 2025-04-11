import os
import sys
import time
import threading
import json
import speech_recognition as sr
import pyautogui
import pyttsx3
import keyboard
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageDraw
import customtkinter as ctk
import webbrowser
import win32gui
import win32con
import win32process
import psutil
import re
import subprocess
import winreg

class VoiceAssistant:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = 4000
        self.recognizer.dynamic_energy_threshold = True
        
        # Configuration
        self.config = {
            "wake_word": "assistant",
            "voice_speed": 150,
            "voice_volume": 1.0,
            "listen_timeout": 5,
            "default_browser": "chrome",
            "search_engine": "google",
            "commands": {
                # Existing commands
                "click": self.click,
                "double click": self.double_click,
                "right click": self.right_click,
                "scroll up": self.scroll_up,
                "scroll down": self.scroll_down,
                "move to": self.move_to,
                "press": self.press_key,
                "type": self.type_text,
                "open": self.open_app,
                "close": self.close_app,
                "minimize": self.minimize_window,
                "maximize": self.maximize_window,
                "minimize all": self.minimize_all_windows,
                "restore all": self.restore_all_windows,
                "minimize all applications": self.minimize_all_applications,
                "maximize all applications": self.maximize_all_applications,
                "stop listening": self.stop_listening,
                "exit": self.exit_app,
                
                # New browser control commands
                "browse to": self.browse_to,
                "search for": self.search_google,
                "open new tab": self.open_new_tab,
                "close tab": self.close_tab,
                "next tab": self.next_tab,
                "previous tab": self.previous_tab,
                "refresh page": self.refresh_page,
                "go back": self.browser_back,
                "go forward": self.browser_forward,
                "scroll to top": self.scroll_to_top,
                "scroll to bottom": self.scroll_to_bottom,
                "zoom in": self.zoom_in,
                "zoom out": self.zoom_out,
                "reset zoom": self.reset_zoom,
                "save page": self.save_page,
                "print page": self.print_page,
                "bookmark page": self.bookmark_page,
                
                # Microsoft Office control commands
                "open word": self.open_word,
                "open excel": self.open_excel,
                "open powerpoint": self.open_powerpoint,
                "open outlook": self.open_outlook,
                "open onenote": self.open_onenote,
                "save document": self.save_document,
                "new document": self.new_document,
                "new spreadsheet": self.new_spreadsheet,
                "new presentation": self.new_presentation,
                "new email": self.new_email,
                "bold text": self.bold_text,
                "italic text": self.italic_text,
                "underline text": self.underline_text,
                "select all": self.select_all,
                "copy": self.copy_text,
                "cut": self.cut_text,
                "paste": self.paste_text,
                "undo": self.undo_action,
                "redo": self.redo_action,
                "insert table": self.insert_table,
                "add slide": self.add_slide,
                "insert image": self.insert_image,
                "switch to reading view": self.reading_view,
                "switch to editing view": self.editing_view,
                "show spelling check": self.spelling_check,
            }
        }
        
        self.running = False
        self.listening = False
        self.gui = None
        
        # Load voices
        voices = self.engine.getProperty('voices')
        self.available_voices = voices
        self.engine.setProperty('voice', voices[0].id)
        self.engine.setProperty('rate', self.config["voice_speed"])
        self.engine.setProperty('volume', self.config["volume"] if "volume" in self.config else 1.0)

    def speak(self, text):
        self.engine.say(text)
        self.engine.runAndWait()
        
    def listen(self):
        with sr.Microphone() as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
            self.gui.update_status("Listening...")
            try:
                audio = self.recognizer.listen(source, timeout=self.config["listen_timeout"])
                self.gui.update_status("Processing...")
                try:
                    text = self.recognizer.recognize_google(audio).lower()
                    self.gui.update_status(f"Heard: {text}")
                    return text
                except sr.UnknownValueError:
                    self.gui.update_status("Could not understand audio")
                    return ""
                except sr.RequestError:
                    self.gui.update_status("Could not request results")
                    return ""
            except sr.WaitTimeoutError:
                self.gui.update_status("Listening timed out")
                return ""
                
    def process_command(self, command):
        if not command:
            return
            
        # Check for wake word
        if self.config["wake_word"] in command:
            self.speak("Yes, I'm listening")
            self.gui.update_status("Activated! Listening for command...")
            command = self.listen()
            
        # Process commands
        for cmd, func in self.config["commands"].items():
            if cmd in command:
                self.gui.update_status(f"Executing: {cmd}")
                params = command.replace(cmd, "").strip()
                func(params)
                return
                
        self.gui.update_status("Command not recognized")
                
    
    def run(self):
        self.running = True
        self.listening = True
        self.speak("Hey BitByBit")
        self.speak("I am your Bitly the voice assistant. How can I help you?")
        self.gui.update_status("Ready to listen")
        
        
        while self.running and self.listening:
            command = self.listen()
            self.process_command(command)
            
    
    def stop_listening(self, params):
        self.listening = False
        self.speak("Voice recognition paused")
        self.gui.update_status("Paused - Click 'Start Listening' to resume")
        
    def exit_app(self, params):
        self.running = False
        self.listening = False
        self.speak("Shutting down")
        self.gui.update_status("Shutting down...")
        time.sleep(1)
        sys.exit(0)
        
    # Existing mouse control commands
    def click(self, params):
        pyautogui.click()
        
    def double_click(self, params):
        pyautogui.doubleClick()
        
    def right_click(self, params):
        pyautogui.rightClick()
        
    def scroll_up(self, params):
        pyautogui.scroll(500)
        
    def scroll_down(self, params):
        pyautogui.scroll(-500)
        
    def move_to(self, params):
        try:
            # Extract coordinates if provided
            coords = params.split()
            if len(coords) >= 2:
                x, y = int(coords[0]), int(coords[1])
                pyautogui.moveTo(x, y, duration=0.5)
            else:
                self.speak("Please specify x and y coordinates")
        except ValueError:
            self.speak("Invalid coordinates")
            
    # Existing keyboard control commands
    def press_key(self, params):
        try:
            key = params.strip()
            if key:
                pyautogui.press(key)
            else:
                self.speak("Please specify a key to press")
        except Exception as e:
            self.speak(f"Error pressing key: {e}")
            
    def type_text(self, params):
        try:
            text = params.strip()
            if text:
                pyautogui.write(text)
            else:
                self.speak("Please specify text to type")
        except Exception as e:
            self.speak(f"Error typing text: {e}")
            
    # Existing application control
    def open_app(self, params):
        app = params.strip()
        try:
            if "chrome" in app or "browser" in app:
                os.system("start chrome")
            elif "edge" in app:
                os.system("start msedge")
            elif "notepad" in app:
                os.system("start notepad")
            elif "explorer" in app or "file" in app:
                os.system("start explorer")
            elif "calculator" in app:
                os.system("start calc")
            else:
                os.system(f"start {app}")
        except Exception as e:
            self.speak(f"Error opening application: {e}")
            
    def close_app(self, params):
        app = params.strip()
        try:
            if app:
                os.system(f"taskkill /f /im {app}.exe")
            else:
                pyautogui.hotkey('alt', 'f4')
        except Exception as e:
            self.speak(f"Error closing application: {e}")
    
    # Existing window control commands
    def minimize_window(self, params):
        try:
            # Get the foreground window
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                # Minimize the window
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                self.speak("Window minimized")
            else:
                self.speak("No active window to minimize")
        except Exception as e:
            self.speak(f"Error minimizing window: {e}")
    
    def maximize_window(self, params):
        try:
            # Get the foreground window
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                # Maximize the window
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                self.speak("Window maximized")
            else:
                self.speak("No active window to maximize")
        except Exception as e:
            self.speak(f"Error maximizing window: {e}")
    
    def minimize_all_windows(self, params):
        try:
            # Minimize all windows (show desktop)
            pyautogui.hotkey('win', 'd')
            self.speak("All windows minimized")
        except Exception as e:
            self.speak(f"Error minimizing all windows: {e}")
    
    def restore_all_windows(self, params):
        try:
            # Restore all windows (press Win+D again)
            pyautogui.hotkey('win', 'd')
            self.speak("Windows restored")
        except Exception as e:
            self.speak(f"Error restoring windows: {e}")
            
    def minimize_all_applications(self, params):
        try:
            # Get all visible windows
            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    windows.append(hwnd)
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            # Minimize each window
            minimized_count = 0
            for hwnd in windows:
                try:
                    # Skip certain system windows
                    if win32gui.GetWindowText(hwnd) and not win32gui.GetWindowText(hwnd).startswith("Voice Control"):
                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                        minimized_count += 1
                except:
                    pass
            
            self.speak(f"Minimized {minimized_count} applications")
            self.gui.log(f"Minimized {minimized_count} applications")
        except Exception as e:
            self.speak(f"Error minimizing applications: {e}")
            self.gui.log(f"Error minimizing applications: {str(e)}")
    
    def maximize_all_applications(self, params):
        try:
            # Get all minimized windows
            def enum_windows_callback(hwnd, windows):
                if win32gui.IsIconic(hwnd) and win32gui.GetWindowText(hwnd):
                    windows.append(hwnd)
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            # Maximize each window
            maximized_count = 0
            for hwnd in windows:
                try:
                    # Skip certain system windows
                    if win32gui.GetWindowText(hwnd) and not win32gui.GetWindowText(hwnd).startswith("Voice Control"):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                        maximized_count += 1
                except:
                    pass
            
            self.speak(f"Maximized {maximized_count} applications")
            self.gui.log(f"Maximized {maximized_count} applications")
        except Exception as e:
            self.speak(f"Error maximizing applications: {e}")
            self.gui.log(f"Error maximizing applications: {str(e)}")
            
    # New browser control functions
    def get_default_browser_path(self):
        try:
            # Try to get the default browser from Windows registry
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                              r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as key:
                program_id = winreg.QueryValueEx(key, "ProgId")[0]
            
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 
                              f"{program_id}\\shell\\open\\command") as key:
                command = winreg.QueryValueEx(key, "")[0].lower()
                
            if "chrome" in command:
                return "chrome"
            elif "firefox" in command:
                return "firefox"
            elif "msedge" in command:
                return "msedge"
            elif "safari" in command:
                return "safari"
            elif "opera" in command:
                return "opera"
            else:
                return "chrome"  # Default to Chrome if can't determine
        except:
            return "chrome"  # Default to Chrome if registry access fails
    
    def browse_to(self, params):
        url = params.strip()
        try:
            # Add http:// if not present and not a domain with common TLD
            if not url.startswith(('http://', 'https://')):
                if not any(url.endswith(tld) for tld in ['.com', '.org', '.net', '.edu', '.gov', '.io']):
                    url = "https://" + url
                else:
                    url = "https://" + url
                    
            self.gui.log(f"Opening URL: {url}")
            webbrowser.open(url)
            self.speak(f"Opening {url}")
        except Exception as e:
            self.speak(f"Error opening URL: {e}")
            self.gui.log(f"Error opening URL: {str(e)}")
    
    def search_google(self, params):
        query = params.strip()
        try:
            if query:
                search_engine = self.config.get("search_engine", "google").lower()
                
                if search_engine == "google":
                    url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
                elif search_engine == "bing":
                    url = f"https://www.bing.com/search?q={query.replace(' ', '+')}"
                elif search_engine == "duckduckgo":
                    url = f"https://duckduckgo.com/?q={query.replace(' ', '+')}"
                elif search_engine == "yahoo":
                    url = f"https://search.yahoo.com/search?p={query.replace(' ', '+')}"
                else:
                    url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
                
                webbrowser.open(url)
                self.speak(f"Searching for {query}")
                self.gui.log(f"Searching for: {query}")
            else:
                self.speak("Please specify what to search for")
        except Exception as e:
            self.speak(f"Error performing search: {e}")
            self.gui.log(f"Error performing search: {str(e)}")
    
    def open_new_tab(self, params):
        try:
            pyautogui.hotkey('ctrl', 't')
            self.speak("New tab opened")
        except Exception as e:
            self.speak(f"Error opening new tab: {e}")
    
    def close_tab(self, params):
        try:
            pyautogui.hotkey('ctrl', 'w')
            self.speak("Tab closed")
        except Exception as e:
            self.speak(f"Error closing tab: {e}")
    
    def next_tab(self, params):
        try:
            pyautogui.hotkey('ctrl', 'tab')
            self.speak("Switched to next tab")
        except Exception as e:
            self.speak(f"Error switching tab: {e}")
    
    def previous_tab(self, params):
        try:
            pyautogui.hotkey('ctrl', 'shift', 'tab')
            self.speak("Switched to previous tab")
        except Exception as e:
            self.speak(f"Error switching tab: {e}")
            
    def refresh_page(self, params):
        try:
            pyautogui.hotkey('f5')
            self.speak("Page refreshed")
        except Exception as e:
            self.speak(f"Error refreshing page: {e}")
    
    def browser_back(self, params):
        try:
            pyautogui.hotkey('alt', 'left')
            self.speak("Going back")
        except Exception as e:
            self.speak(f"Error going back: {e}")
    
    def browser_forward(self, params):
        try:
            pyautogui.hotkey('alt', 'right')
            self.speak("Going forward")
        except Exception as e:
            self.speak(f"Error going forward: {e}")
    
    def scroll_to_top(self, params):
        try:
            pyautogui.hotkey('home')
            self.speak("Scrolled to top")
        except Exception as e:
            self.speak(f"Error scrolling: {e}")
    
    def scroll_to_bottom(self, params):
        try:
            pyautogui.hotkey('end')
            self.speak("Scrolled to bottom")
        except Exception as e:
            self.speak(f"Error scrolling: {e}")
    
    def zoom_in(self, params):
        try:
            pyautogui.hotkey('ctrl', '+')
            self.speak("Zoomed in")
        except Exception as e:
            self.speak(f"Error zooming: {e}")
    
    def zoom_out(self, params):
        try:
            pyautogui.hotkey('ctrl', '-')
            self.speak("Zoomed out")
        except Exception as e:
            self.speak(f"Error zooming: {e}")
    
    def reset_zoom(self, params):
        try:
            pyautogui.hotkey('ctrl', '0')
            self.speak("Zoom reset")
        except Exception as e:
            self.speak(f"Error resetting zoom: {e}")
    
    def save_page(self, params):
        try:
            pyautogui.hotkey('ctrl', 's')
            self.speak("Saving page")
        except Exception as e:
            self.speak(f"Error saving page: {e}")
    
    def print_page(self, params):
        try:
            pyautogui.hotkey('ctrl', 'p')
            self.speak("Opening print dialog")
        except Exception as e:
            self.speak(f"Error printing page: {e}")
    
    def bookmark_page(self, params):
        try:
            pyautogui.hotkey('ctrl', 'd')
            self.speak("Bookmarking page")
        except Exception as e:
            self.speak(f"Error bookmarking page: {e}")
    
    # Microsoft Office application control functions
    def get_office_path(self):
        # Try common installation paths
        office_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16",
            r"C:\Program Files\Microsoft Office\Office16",
            r"C:\Program Files (x86)\Microsoft Office\Office16",
            r"C:\Program Files\Microsoft Office\Office15",
            r"C:\Program Files (x86)\Microsoft Office\Office15",
            r"C:\Program Files\Microsoft Office\Office14",
            r"C:\Program Files (x86)\Microsoft Office\Office14",
        ]
        
        for path in office_paths:
            if os.path.exists(path):
                return path
                
        return None
    
    def open_word(self, params):
        try:
            os.system("start winword")
            self.speak("Opening Microsoft Word")
        except Exception as e:
            self.speak(f"Error opening Word: {e}")
    
    def open_excel(self, params):
        try:
            os.system("start excel")
            self.speak("Opening Microsoft Excel")
        except Exception as e:
            self.speak(f"Error opening Excel: {e}")
    
    def open_powerpoint(self, params):
        try:
            os.system("start powerpnt")
            self.speak("Opening Microsoft PowerPoint")
        except Exception as e:
            self.speak(f"Error opening PowerPoint: {e}")
    
    def open_outlook(self, params):
        try:
            os.system("start outlook")
            self.speak("Opening Microsoft Outlook")
        except Exception as e:
            self.speak(f"Error opening Outlook: {e}")
    
    def open_onenote(self, params):
        try:
            os.system("start onenote")
            self.speak("Opening Microsoft OneNote")
        except Exception as e:
            self.speak(f"Error opening OneNote: {e}")
    
    def save_document(self, params):
        try:
            pyautogui.hotkey('ctrl', 's')
            self.speak("Saving document")
        except Exception as e:
            self.speak(f"Error saving document: {e}")
    
    def new_document(self, params):
        try:
            pyautogui.hotkey('ctrl', 'n')
            self.speak("Creating new document")
        except Exception as e:
            self.speak(f"Error creating new document: {e}")
    
    def new_spreadsheet(self, params):
        try:
            self.open_excel("")
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'n')
            self.speak("Creating new spreadsheet")
        except Exception as e:
            self.speak(f"Error creating new spreadsheet: {e}")
    
    def new_presentation(self, params):
        try:
            self.open_powerpoint("")
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'n')
            self.speak("Creating new presentation")
        except Exception as e:
            self.speak(f"Error creating new presentation: {e}")
    
    def new_email(self, params):
        try:
            # Check if Outlook is running
            outlook_running = False
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'outlook' in proc.info['name'].lower():
                    outlook_running = True
                    break
            
            if not outlook_running:
                self.open_outlook("")
                time.sleep(2)
            
            pyautogui.hotkey('ctrl', 'n')
            self.speak("Creating new email")
        except Exception as e:
            self.speak(f"Error creating new email: {e}")
    
    def bold_text(self, params):
        try:
            pyautogui.hotkey('ctrl', 'b')
            self.speak("Bold text applied")
        except Exception as e:
            self.speak(f"Error applying bold: {e}")
    
    def italic_text(self, params):
        try:
            pyautogui.hotkey('ctrl', 'i')
            self.speak("Italic text applied")
        except Exception as e:
            self.speak(f"Error applying italic: {e}")
    
    def underline_text(self, params):
        try:
            pyautogui.hotkey('ctrl', 'u')
            self.speak("Underline text applied")
        except Exception as e:
            self.speak(f"Error applying underline: {e}")
    
    def select_all(self, params):
        try:
            pyautogui.hotkey('ctrl', 'a')
            self.speak("Selected all content")
        except Exception as e:
            self.speak(f"Error selecting all: {e}")
    
    def copy_text(self, params):
        try:
            pyautogui.hotkey('ctrl', 'c')
            self.speak("Copied to clipboard")
        except Exception as e:
            self.speak(f"Error copying: {e}")
    
    def cut_text(self, params):
        try:
            pyautogui.hotkey('ctrl', 'x')
            self.speak("Cut to clipboard")
        except Exception as e:
            self.speak(f"Error cutting: {e}")
    
    def paste_text(self, params):
        try:
            pyautogui.hotkey('ctrl', 'v')
            self.speak("Pasted from clipboard")
        except Exception as e:
            self.speak(f"Error pasting: {e}")
    
    def undo_action(self, params):
        try:
            pyautogui.hotkey('ctrl', 'z')
            self.speak("Undoing last action")
        except Exception as e:
            self.speak(f"Error undoing: {e}")
    
    def redo_action(self, params):
        try:
            pyautogui.hotkey('ctrl', 'y')
            self.speak("Redoing action")
        except Exception as e:
            self.speak(f"Error redoing: {e}")
    
    def insert_table(self, params):
        try:
            # For Word
            active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
            
            if "word" in active_window:
                # Word specific table insertion
                pyautogui.hotkey('alt')
                time.sleep(0.3)
                pyautogui.press('n')
                time.sleep(0.3)
                pyautogui.press('t')
                self.speak("Inserting table in Word")
            elif "excel" in active_window:
                self.speak("In Excel, just select cells to create a table")
            elif "powerpoint" in active_window:
                pyautogui.hotkey('alt')
                time.sleep(0.3)
                pyautogui.press('n')
                time.sleep(0.3)
                pyautogui.press('t')
                self.speak("Inserting table in PowerPoint")
            else:
                self.speak("Please open a Microsoft Office application first")
        except Exception as e:
            self.speak(f"Error inserting table: {e}")
    
    def add_slide(self, params):
        try:
            active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
            
            if "powerpoint" in active_window:
                pyautogui.hotkey('ctrl', 'm')
                self.speak("Adding new slide")
            else:
                self.speak("Please open PowerPoint first")
        except Exception as e:
            self.speak(f"Error adding slide: {e}")
    
    def insert_image(self, params):
        try:
            active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
            
            if any(app in active_window for app in ["word", "powerpoint", "excel", "outlook"]):
                pyautogui.hotkey('alt')
                time.sleep(0.3)
                pyautogui.press('n')
                time.sleep(0.3)
                pyautogui.press('p')
                self.speak("Opening insert picture dialog")
            else:
                self.speak("Please open a Microsoft Office application first")
        except Exception as e:
            self.speak(f"Error inserting image: {e}")
    
    def reading_view(self, params):
        try:
            active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
            
            if "word" in active_window:
                pyautogui.hotkey('alt', 'w', 'f')
                self.speak("Switching to reading view")
            elif "powerpoint" in active_window:
                pyautogui.press('f5')
                self.speak("Starting slideshow")
            else:
                self.speak("Please open Word or PowerPoint first")
        except Exception as e:
            self.speak(f"Error switching view: {e}")
    
    def editing_view(self, params):
        try:
            active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
            
            if "word" in active_window:
                pyautogui.hotkey('esc')
                self.speak("Switching to editing view")
            elif "powerpoint" in active_window:
                pyautogui.hotkey('esc')
                self.speak("Exiting slideshow")
            else:
                self.speak("Please open Word or PowerPoint first")
        except Exception as e:
            self.speak(f"Error switching view: {e}")
    
    def spelling_check(self, params):
        try:
            pyautogui.hotkey('f7')
            self.speak("Opening spelling check")
        except Exception as e:
            self.speak(f"Error opening spelling check: {e}")

class ModernGUI:
    def __init__(self, assistant):
        self.assistant = assistant
        self.assistant.gui = self
        
        # Set appearance mode and default color theme
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        # Create the main window
        self.root = ctk.CTk()
        self.root.title("Voice Control Assistant")
        self.root.geometry("900x600")
        self.root.minsize(900, 600)
        
        # Create a custom icon
        self.root.iconbitmap(default="icon.ico") if os.path.exists("icon.ico") else None
        
        # Configure grid layout (continued)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=3)
        self.root.grid_rowconfigure(0, weight=1)
        
        # Create sidebar frame
        self.sidebar_frame = ctk.CTkFrame(self.root, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)
        
        # App logo/name
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Voice Control", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Control buttons in sidebar
        self.start_button = ctk.CTkButton(self.sidebar_frame, text="Start Listening", command=self.start_listening)
        self.start_button.grid(row=1, column=0, padx=20, pady=10)
        
        self.stop_button = ctk.CTkButton(self.sidebar_frame, text="Stop Listening", command=self.stop_listening)
        self.stop_button.grid(row=2, column=0, padx=20, pady=10)
        
        self.settings_button = ctk.CTkButton(self.sidebar_frame, text="Settings", command=self.open_settings)
        self.settings_button.grid(row=3, column=0, padx=20, pady=10)
        
        self.help_button = ctk.CTkButton(self.sidebar_frame, text="Help", command=self.show_help)
        self.help_button.grid(row=4, column=0, padx=20, pady=10)
        
        self.exit_button = ctk.CTkButton(self.sidebar_frame, text="Exit", fg_color="darkred", hover_color="red", command=self.exit_app)
        self.exit_button.grid(row=5, column=0, padx=20, pady=10)
        
        # Appearance mode option
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_option = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                   command=self.change_appearance_mode)
        self.appearance_mode_option.grid(row=7, column=0, padx=20, pady=(10, 20))
        
        # Create main content frame
        self.content_frame = ctk.CTkFrame(self.root)
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=0)
        self.content_frame.grid_rowconfigure(1, weight=3)
        self.content_frame.grid_rowconfigure(2, weight=1)
        
        # Status display
        self.status_frame = ctk.CTkFrame(self.content_frame)
        self.status_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        
        self.status_label = ctk.CTkLabel(self.status_frame, text="Status:", font=ctk.CTkFont(weight="bold"))
        self.status_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.status_text = ctk.CTkLabel(self.status_frame, text="Not listening", font=ctk.CTkFont(size=16))
        self.status_text.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Log display
        self.log_frame = ctk.CTkFrame(self.content_frame)
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(0, weight=0)
        self.log_frame.grid_rowconfigure(1, weight=1)
        
        self.log_label = ctk.CTkLabel(self.log_frame, text="Command Log", font=ctk.CTkFont(weight="bold"))
        self.log_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.log_text = ctk.CTkTextbox(self.log_frame, width=400, height=300, font=ctk.CTkFont(size=14))
        self.log_text.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.log_text.configure(state="disabled")
        
        # Command panel with tabs for different categories
        self.command_frame = ctk.CTkFrame(self.content_frame)
        self.command_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.command_frame.grid_columnconfigure(0, weight=1)
        self.command_frame.grid_rowconfigure(0, weight=0)
        self.command_frame.grid_rowconfigure(1, weight=1)
        
        self.command_label = ctk.CTkLabel(self.command_frame, text="Available Commands", font=ctk.CTkFont(weight="bold"))
        self.command_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        # Create tabview for organizing commands by category
        self.command_tabs = ctk.CTkTabview(self.command_frame)
        self.command_tabs.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        
        # Create tabs for different command categories
        self.tab_basic = self.command_tabs.add("Basic")
        self.tab_browser = self.command_tabs.add("Browser")
        self.tab_office = self.command_tabs.add("Office")
        
        # Add commands to basic tab
        basic_commands = [
            "click", "double click", "right click", "scroll up", "scroll down",
            "move to", "press", "type", "open", "close", "minimize", "maximize", 
            "minimize all", "restore all", "stop listening", "exit","browse to", "search for", "open new tab", "close tab", "next tab", 
            "previous tab", "refresh page", "go back", "go forward", "scroll to top",
            "scroll to bottom", "zoom in", "zoom out", "reset zoom", "save page",
            "print page", "bookmark page", "open word", "open excel", "open powerpoint", "open outlook", "open onenote",
            "save document", "new document", "new spreadsheet", "new presentation", "new email",
            "bold text", "italic text", "underline text", "select all", "copy", 
            "cut", "paste", "undo", "redo", "insert table",
            "add slide", "insert image", "switch to reading view", "switch to editing view", "show spelling check"
        ]
        
        self.setup_command_buttons(self.tab_basic, basic_commands, rows=5, cols=3)
        
        # Add commands to browser tab
        browser_commands = [
            "click", "double click", "right click", "scroll up", "scroll down",
            "move to", "press", "type", "open", "close", "minimize", "maximize", 
            "minimize all", "restore all", "stop listening", "exit","browse to", "search for", "open new tab", "close tab", "next tab", 
            "previous tab", "refresh page", "go back", "go forward", "scroll to top",
            "scroll to bottom", "zoom in", "zoom out", "reset zoom", "save page",
            "print page", "bookmark page", "open word", "open excel", "open powerpoint", "open outlook", "open onenote",
            "save document", "new document", "new spreadsheet", "new presentation", "new email",
            "bold text", "italic text", "underline text", "select all", "copy", 
            "cut", "paste", "undo", "redo", "insert table",
            "add slide", "insert image", "switch to reading view", "switch to editing view", "show spelling check"
        ]
        
        self.setup_command_buttons(self.tab_browser, browser_commands, rows=6, cols=3)
        
        # Add commands to office tab
        office_commands = [
                        "click", "double click", "right click", "scroll up", "scroll down",
            "move to", "press", "type", "open", "close", "minimize", "maximize", 
            "minimize all", "restore all", "stop listening", "exit","browse to", "search for", "open new tab", "close tab", "next tab", 
            "previous tab", "refresh page", "go back", "go forward", "scroll to top",
            "scroll to bottom", "zoom in", "zoom out", "reset zoom", "save page",
            "print page", "bookmark page", "open word", "open excel", "open powerpoint", "open outlook", "open onenote",
            "save document", "new document", "new spreadsheet", "new presentation", "new email",
            "bold text", "italic text", "underline text", "select all", "copy", 
            "cut", "paste", "undo", "redo", "insert table",
            "add slide", "insert image", "switch to reading view", "switch to editing view", "show spelling check"
        ]
        
        self.setup_command_buttons(self.tab_office, office_commands, rows=9, cols=3)
            
        # Set initial values
        self.appearance_mode_option.set("System")
        self.update_status("Ready")
        self.log("Enhanced Voice Control Assistant initialized. Click 'Start Listening' to begin.")
        
        # Thread for voice assistant
        self.assistant_thread = None
        
    def setup_command_buttons(self, parent, commands, rows, cols):
        """Helper method to set up command buttons in a grid layout"""
        for i, cmd in enumerate(commands):
            row, col = divmod(i, cols)
            if row < rows:  # Limit number of rows to display
                btn = ctk.CTkButton(parent, text=cmd, 
                                command=lambda c=cmd: self.display_command_info(c),
                                width=120, height=30)
                btn.grid(row=row, column=col, padx=5, pady=5)
        
    def start(self):
        self.root.mainloop()
        
    def start_listening(self):
        if self.assistant_thread and self.assistant_thread.is_alive():
            messagebox.showinfo("Already Running", "Voice assistant is already listening!")
            return
            
        self.assistant.listening = True
        self.assistant_thread = threading.Thread(target=self.assistant.run)
        self.assistant_thread.daemon = True
        self.assistant_thread.start()
        self.update_status("Listening...")
        self.log("Voice assistant started. Say the wake word to activate.")
        
    def stop_listening(self):
        if self.assistant_thread and self.assistant_thread.is_alive():
            self.assistant.stop_listening("")
            self.update_status("Stopped")
            self.log("Voice assistant stopped.")
        else:
            messagebox.showinfo("Not Running", "Voice assistant is not currently listening!")
            
    def exit_app(self):
        if messagebox.askyesno("Exit Application", "Are you sure you want to exit?"):
            self.assistant.running = False
            self.assistant.listening = False
            self.root.quit()
            sys.exit(0)
            
    def update_status(self, status):
        self.status_text.configure(text=status)
        self.root.update_idletasks()
        
    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        
    def change_appearance_mode(self, new_appearance_mode):
        ctk.set_appearance_mode(new_appearance_mode)
        
    def open_settings(self):
        settings_window = ctk.CTkToplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("500x550")
        settings_window.transient(self.root)
        settings_window.focus_set()
        settings_window.grab_set()
        
        settings_window.grid_columnconfigure(0, weight=1)
        settings_window.grid_columnconfigure(1, weight=2)
        
        # Create tabview for settings
        settings_tabs = ctk.CTkTabview(settings_window)
        settings_tabs.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")
        settings_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        tab_general = settings_tabs.add("General")
        tab_browser = settings_tabs.add("Browser")
        tab_voice = settings_tabs.add("Voice")
        
        # General settings tab
        tab_general.grid_columnconfigure(0, weight=1)
        tab_general.grid_columnconfigure(1, weight=2)
        
        # Wake word setting
        wake_label = ctk.CTkLabel(tab_general, text="Wake Word:")
        wake_label.grid(row=0, column=0, padx=20, pady=20, sticky="e")
        
        wake_entry = ctk.CTkEntry(tab_general, width=200)
        wake_entry.grid(row=0, column=1, padx=20, pady=20, sticky="w")
        wake_entry.insert(0, self.assistant.config["wake_word"])
        
        # Listen timeout setting
        timeout_label = ctk.CTkLabel(tab_general, text="Listen Timeout (sec):")
        timeout_label.grid(row=1, column=0, padx=20, pady=20, sticky="e")
        
        timeout_entry = ctk.CTkEntry(tab_general, width=200)
        timeout_entry.grid(row=1, column=1, padx=20, pady=20, sticky="w")
        timeout_entry.insert(0, str(self.assistant.config["listen_timeout"]))
        
        # Browser settings tab
        tab_browser.grid_columnconfigure(0, weight=1)
        tab_browser.grid_columnconfigure(1, weight=2)
        
        # Default browser setting
        browser_label = ctk.CTkLabel(tab_browser, text="Default Browser:")
        browser_label.grid(row=0, column=0, padx=20, pady=20, sticky="e")
        
        browser_option = ctk.CTkOptionMenu(tab_browser, values=["chrome", "firefox", "edge", "system default"])
        browser_option.grid(row=0, column=1, padx=20, pady=20, sticky="w")
        browser_option.set(self.assistant.config.get("default_browser", "chrome"))
        
        # Search engine setting
        search_label = ctk.CTkLabel(tab_browser, text="Search Engine:")
        search_label.grid(row=1, column=0, padx=20, pady=20, sticky="e")
        
        search_option = ctk.CTkOptionMenu(tab_browser, values=["google", "bing", "duckduckgo", "yahoo"])
        search_option.grid(row=1, column=1, padx=20, pady=20, sticky="w")
        search_option.set(self.assistant.config.get("search_engine", "google"))
        
        # Voice settings tab
        tab_voice.grid_columnconfigure(0, weight=1)
        tab_voice.grid_columnconfigure(1, weight=2)
        
        # Voice selection
        voice_label = ctk.CTkLabel(tab_voice, text="Voice:")
        voice_label.grid(row=0, column=0, padx=20, pady=20, sticky="e")
        
        voice_names = [voice.name for voice in self.assistant.available_voices]
        voice_option = ctk.CTkOptionMenu(tab_voice, values=voice_names)
        voice_option.grid(row=0, column=1, padx=20, pady=20, sticky="w")
        voice_option.set(voice_names[0])
        
        # Voice speed setting
        speed_label = ctk.CTkLabel(tab_voice, text="Voice Speed:")
        speed_label.grid(row=1, column=0, padx=20, pady=20, sticky="e")
        
        speed_slider = ctk.CTkSlider(tab_voice, from_=100, to=200)
        speed_slider.grid(row=1, column=1, padx=20, pady=20, sticky="w")
        speed_slider.set(self.assistant.config["voice_speed"])
        
        speed_value = ctk.CTkLabel(tab_voice, text=str(self.assistant.config["voice_speed"]))
        speed_value.grid(row=1, column=1, padx=(200, 20), pady=20, sticky="w")
        
        # Voice volume setting
        volume_label = ctk.CTkLabel(tab_voice, text="Voice Volume:")
        volume_label.grid(row=2, column=0, padx=20, pady=20, sticky="e")
        
        volume_slider = ctk.CTkSlider(tab_voice, from_=0, to=1, number_of_steps=10)
        volume_slider.grid(row=2, column=1, padx=20, pady=20, sticky="w")
        volume_slider.set(self.assistant.config.get("voice_volume", 1.0))
        
        volume_value = ctk.CTkLabel(tab_voice, text=f"{int(self.assistant.config.get('voice_volume', 1.0) * 100)}%")
        volume_value.grid(row=2, column=1, padx=(200, 20), pady=20, sticky="w")
        
        # Update functions
        def update_speed(value):
            speed_value.configure(text=str(int(value)))
            
        def update_volume(value):
            volume_value.configure(text=f"{int(value * 100)}%")
            
        speed_slider.configure(command=update_speed)
        volume_slider.configure(command=update_volume)
        
        # Save button
        def save_settings():
            # Save general settings
            self.assistant.config["wake_word"] = wake_entry.get()
            self.assistant.config["listen_timeout"] = int(timeout_entry.get())
            
            # Save browser settings
            self.assistant.config["default_browser"] = browser_option.get()
            self.assistant.config["search_engine"] = search_option.get()
            
            # Save voice settings
            self.assistant.config["voice_speed"] = int(speed_slider.get())
            self.assistant.config["voice_volume"] = float(volume_slider.get())
            
            # Set voice
            voice_index = voice_names.index(voice_option.get())
            self.assistant.engine.setProperty('voice', self.assistant.available_voices[voice_index].id)
            self.assistant.engine.setProperty('rate', self.assistant.config["voice_speed"])
            self.assistant.engine.setProperty('volume', self.assistant.config["voice_volume"])
            
            self.log(f"Settings updated: Wake word = '{self.assistant.config['wake_word']}', Voice speed = {self.assistant.config['voice_speed']}")
            settings_window.destroy()
            
        save_button = ctk.CTkButton(settings_window, text="Save", command=save_settings)
        save_button.grid(row=1, column=0, columnspan=2, padx=20, pady=20)
        
    def show_help(self):
        help_window = ctk.CTkToplevel(self.root)
        help_window.title("Help")
        help_window.geometry("800x600")
        help_window.transient(self.root)
        help_window.focus_set()
        
        help_window.grid_columnconfigure(0, weight=1)
        help_window.grid_rowconfigure(0, weight=0)
        help_window.grid_rowconfigure(1, weight=1)
        
        # Title
        title_label = ctk.CTkLabel(help_window, text="Voice Control Assistant Help", 
                                 font=ctk.CTkFont(size=20, weight="bold"))
        title_label.grid(row=0, column=0, padx=20, pady=20)
        
        # Create tabview for help content
        help_tabs = ctk.CTkTabview(help_window)
        help_tabs.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        
        # Create tabs
        tab_general = help_tabs.add("General")
        tab_browser = help_tabs.add("Browser")
        tab_office = help_tabs.add("Office")
        
        # General help content
        general_text = ctk.CTkTextbox(tab_general, width=700, height=400)
        general_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        general_content = """
# Voice Control Assistant Help

## Getting Started
1. Click 'Start Listening' to activate the voice assistant
2. Say the wake word (default: "assistant") to begin giving commands
3. Speak your command clearly

## Basic Commands

### Mouse Control
- "click" - Perform a left mouse click
- "double click" - Perform a double click
- "right click" - Perform a right click
- "scroll up" - Scroll up
- "scroll down" - Scroll down
- "move to X Y" - Move mouse to coordinates X,Y

### Keyboard Control
- "press [key]" - Press a keyboard key (e.g., "press enter", "press space")
- "type [text]" - Type the specified text

### Window Control
- "minimize" - Minimize the current active window
- "maximize" - Maximize the current active window
- "minimize all" - Minimize all windows (show desktop)
- "restore all" - Restore all minimized windows
- "minimize all applications" - Minimize all running applications
- "maximize all applications" - Maximize all minimized applications

### Application Control
- "open [app]" - Open an application (e.g., "open chrome", "open notepad")
- "close [app]" - Close an application (or use without an app name to close current window)

### Assistant Control
- "stop listening" - Pause voice recognition
- "exit" - Exit the application

## Tips
- Speak clearly and in a moderate pace
- Make sure your microphone is working properly
- Adjust settings to optimize recognition for your voice
- For best results, use in a quiet environment
        """
        
        general_text.insert("1.0", general_content)
        general_text.configure(state="disabled")
        
        # Browser help content
        browser_text = ctk.CTkTextbox(tab_browser, width=700, height=400)
        browser_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        browser_content = """
# Browser Control Commands

### Navigation
- "browse to [website]" - Open a specific website (e.g., "browse to example.com")
- "search for [query]" - Search using the default search engine
- "go back" - Navigate back to the previous page
- "go forward" - Navigate forward to the next page
- "refresh page" - Refresh the current page

### Tab Management
- "open new tab" - Open a new browser tab
- "close tab" - Close the current tab
- "next tab" - Switch to the next tab
- "previous tab" - Switch to the previous tab

### Page Interaction
- "scroll to top" - Scroll to the top of the page
- "scroll to bottom" - Scroll to the bottom of the page
- "zoom in" - Increase the page zoom
- "zoom out" - Decrease the page zoom
- "reset zoom" - Reset the page zoom to default
- "save page" - Open the save page dialog
- "print page" - Open the print dialog
- "bookmark page" - Bookmark the current page

## Tips
- For the "browse to" command, you can include "http://" or just the domain name
- The default search engine can be changed in the settings
- Some commands may work differently depending on the browser being used
- For most reliable results, use Chrome or Edge browsers
        """
        
        browser_text.insert("1.0", browser_content)
        browser_text.configure(state="disabled")
        
        # Office help content
        office_text = ctk.CTkTextbox(tab_office, width=700, height=400)
        office_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        office_content = """
# Microsoft Office Control Commands

### Application Opening
- "open word" - Open Microsoft Word
- "open excel" - Open Microsoft Excel
- "open powerpoint" - Open Microsoft PowerPoint
- "open outlook" - Open Microsoft Outlook
- "open onenote" - Open Microsoft OneNote

### Document Control
- "new document" - Create a new document (works in most Office apps)
- "new spreadsheet" - Create a new Excel spreadsheet
- "new presentation" - Create a new PowerPoint presentation
- "new email" - Create a new email in Outlook
- "save document" - Save the current document

### Text Formatting
- "bold text" - Apply bold formatting to selected text
- "italic text" - Apply italic formatting to selected text
- "underline text" - Apply underline formatting to selected text

### Editing
- "select all" - Select all content in the document
- "copy" - Copy selected content to clipboard
- "cut" - Cut selected content to clipboard
- "paste" - Paste content from clipboard
- "undo" - Undo the last action
- "redo" - Redo the last undone action

### Content Insertion
- "insert table" - Open the insert table dialog
- "insert image" - Open the insert image dialog
- "add slide" - Add a new slide in PowerPoint

### View Control
- "switch to reading view" - Switch to reading view (Word) or start slideshow (PowerPoint)
- "switch to editing view" - Return to editing view
- "show spelling check" - Open the spelling check dialog

## Tips
- Make sure the appropriate Office application is open before using specific commands
- Some commands are application-specific (e.g., "add slide" only works in PowerPoint)
- For text formatting commands, select the text first before giving the command
- Office commands use standard keyboard shortcuts, so they should work with all Office versions
        """
        
        office_text.insert("1.0", office_content)
        office_text.configure(state="disabled")
        
    def display_command_info(self, command):
        # Update with new commands
        info_dict = {
            # Basic commands
            "click": "Performs a left mouse click at the current position.",
            "double click": "Performs a double-click at the current mouse position.",
            "right click": "Performs a right mouse click at the current position.",
            "scroll up": "Scrolls up on the page.",
            "scroll down": "Scrolls down on the page.",
            "move to": "Moves the mouse to specified X Y coordinates. Example: 'move to 500 300'",
            "press": "Presses the specified keyboard key. Example: 'press enter'",
            "type": "Types the specified text. Example: 'type Hello world'",
            "open": "Opens the specified application. Example: 'open chrome'",
            "close": "Closes the current application or specified one. Example: 'close notepad'",
            "minimize": "Minimizes the currently active window.",
            "maximize": "Maximizes the currently active window.",
            "minimize all": "Minimizes all windows (shows desktop).",
            "restore all": "Restores all minimized windows.",
            "minimize all applications": "Minimizes all running applications individually.",
            "maximize all applications": "Maximizes all currently minimized applications.",
            "stop listening": "Pauses voice recognition until manually restarted.",
            "exit": "Exits the voice assistant application.",
            
            # Browser commands
            "browse to": "Opens the specified website. Example: 'browse to example.com'",
            "search for": "Searches using the default search engine. Example: 'search for voice assistants'",
            "open new tab": "Opens a new browser tab.",
            "close tab": "Closes the current browser tab.",
            "next tab": "Switches to the next browser tab.",
            "previous tab": "Switches to the previous browser tab.",
            "refresh page": "Refreshes the current web page.",
            "go back": "Navigates back to the previous page.",
            "go forward": "Navigates forward to the next page.",
            "scroll to top": "Scrolls to the top of the web page.",
            "scroll to bottom": "Scrolls to the bottom of the web page.",
            "zoom in": "Increases the browser zoom level.",
            "zoom out": "Decreases the browser zoom level.",
            "reset zoom": "Resets the browser zoom level to default.",
            "save page": "Opens the save page dialog.",
            "print page": "Opens the print dialog for the current page.",
            "bookmark page": "Bookmarks the current page.",
            
            # Microsoft Office commands
            "open word": "Opens Microsoft Word.",
            "open excel": "Opens Microsoft Excel.",
            "open powerpoint": "Opens Microsoft PowerPoint.",
            "open outlook": "Opens Microsoft Outlook.",
            "open onenote": "Opens Microsoft OneNote.",
            "save document": "Saves the current document.",
            "new document": "Creates a new document.",
            "new spreadsheet": "Creates a new Excel spreadsheet.",
            "new presentation": "Creates a new PowerPoint presentation.",
            "new email": "Creates a new email in Outlook.",
            "bold text": "Applies bold formatting to selected text.",
            "italic text": "Applies italic formatting to selected text.",
            "underline text": "Applies underline formatting to selected text.",
            "select all": "Selects all content in the document.",
            "copy": "Copies selected content to clipboard.",
            "cut": "Cuts selected content to clipboard.",
            "paste": "Pastes content from clipboard.",
            "undo": "Undoes the last action.",
            "redo": "Redoes the last undone action.",
            "insert table": "Opens the insert table dialog.",
            "add slide": "Adds a new slide in PowerPoint.",
            "insert image": "Opens the insert image dialog.",
            "switch to reading view": "Switches to reading view in Word or starts slideshow in PowerPoint.",
            "switch to editing view": "Returns to editing view from reading view or slideshow.",
            "show spelling check": "Opens the spelling checker."
        }
        
        messagebox.showinfo(f"Command: {command}", info_dict.get(command, "No information available"))

def main():
    # Set up PyAutoGUI settings
    pyautogui.FAILSAFE = True
    
    # Create application
    assistant = VoiceAssistant()
    gui = ModernGUI(assistant)
    
    # Start GUI
    gui.start()

if __name__ == "__main__":
    main()
