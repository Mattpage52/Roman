import keyboard
import time
import tkinter as tk
from tkinter import ttk
from configparser import ConfigParser
import os
import win32api
import threading
import sys

class KeyPresserApp:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("Key Presser")
            self.root.geometry("300x200")
            
            # Load config with error handling
            try:
                self.config = ConfigParser()
                if not os.path.exists('config.ini'):
                    self.create_default_config()
                self.config.read('config.ini')
                self.trigger_key = int(self.config['Settings']['trigger_key'])
            except Exception as e:
                print(f"Error loading config: {e}")
                # Set default if config fails
                self.trigger_key = 1
                self.create_default_config()
            
            # Load additional settings from config
            self.repeat_enabled = self.config.getboolean('Settings', 'repeat_enabled', fallback=True)
            self.key_delay = self.config.getfloat('Settings', 'key_delay', fallback=0.05)
            self.speed_multiplier = self.config.getfloat('Settings', 'speed_multiplier', fallback=1.0)
            
            # Create main frame
            self.main_frame = ttk.Frame(root, padding="10")
            self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Status label
            self.status_label = ttk.Label(
                self.main_frame, 
                text=f"Current trigger: Mouse Button {self.trigger_key}"
            )
            self.status_label.grid(row=0, column=0, pady=10)
            
            # Config button
            self.config_button = ttk.Button(
                self.main_frame, 
                text="Press to configure new button", 
                command=self.start_capture
            )
            self.config_button.grid(row=1, column=0, pady=5)
            
            # Start/Stop button
            self.running = False
            self.toggle_button = ttk.Button(
                self.main_frame, 
                text="Start", 
                command=self.toggle_script
            )
            self.toggle_button.grid(row=2, column=0, pady=5)
            
            # Flags
            self.capturing = False
            self.running = False
            self.monitor_thread = None
            
            # Add repeat toggle checkbox
            self.repeat_var = tk.BooleanVar(value=self.repeat_enabled)
            self.repeat_checkbox = ttk.Checkbutton(
                self.main_frame,
                text="Repeat while held",
                variable=self.repeat_var,
                command=self.save_settings
            )
            self.repeat_checkbox.grid(row=3, column=0, pady=5)
            
            # Add delay slider
            self.delay_frame = ttk.Frame(self.main_frame)
            self.delay_frame.grid(row=4, column=0, pady=5)
            ttk.Label(self.delay_frame, text="Key press delay:").grid(row=0, column=0)
            self.delay_var = tk.DoubleVar(value=self.key_delay)
            self.delay_slider = ttk.Scale(
                self.delay_frame,
                from_=0.01,
                to=0.5,
                variable=self.delay_var,
                orient='horizontal',
                command=lambda _: self.save_settings()
            )
            self.delay_slider.grid(row=0, column=1)
            
            # Add speed dropdown
            self.speed_frame = ttk.Frame(self.main_frame)
            self.speed_frame.grid(row=5, column=0, pady=5)
            ttk.Label(self.speed_frame, text="Speed:").grid(row=0, column=0, padx=5)
            
            self.speeds = {
                'Very Slow': 0.3,
                'Slow': 0.6,
                'Normal': 1.0,
                'Fast': 1.5,
                'Very Fast': 2.0
            }
            
            self.speed_var = tk.StringVar(value=[k for k, v in self.speeds.items() if v == self.speed_multiplier][0])
            self.speed_combo = ttk.Combobox(
                self.speed_frame,
                textvariable=self.speed_var,
                values=list(self.speeds.keys()),
                state='readonly',
                width=10
            )
            self.speed_combo.grid(row=0, column=1, padx=5)
            self.speed_combo.bind('<<ComboboxSelected>>', lambda _: self.save_settings())
            
            # Add system-specific delay compensation
            self.system_delay_compensation = 1.0
            try:
                # Test system responsiveness
                start = time.perf_counter()
                time.sleep(0.01)
                end = time.perf_counter()
                actual_delay = end - start
                self.system_delay_compensation = 0.01 / actual_delay
                print(f"System delay compensation: {self.system_delay_compensation}")
            except:
                print("Could not calculate system delay compensation")
            
        except Exception as e:
            print(f"Error during initialization: {e}")
            sys.exit(1)
    
    def create_default_config(self):
        try:
            self.config = ConfigParser()
            self.config['Settings'] = {
                'trigger_key': '1',
                'repeat_enabled': 'true',
                'key_delay': '0.05',
                'speed_multiplier': '1.0'
            }
            with open('config.ini', 'w') as f:
                self.config.write(f)
        except Exception as e:
            print(f"Error creating config: {e}")
    
    def start_capture(self):
        if not self.capturing:
            self.capturing = True
            self.config_button.config(text="Press any mouse button...")
            
            if self.running:
                self.toggle_script()
            
            capture_thread = threading.Thread(target=self.capture_mouse_button)
            capture_thread.daemon = True
            capture_thread.start()
    
    def capture_mouse_button(self):
        try:
            previous_states = [False] * 32
            
            while self.capturing:
                try:
                    current_states = [win32api.GetAsyncKeyState(i) < 0 for i in range(1, 33)]
                    
                    for i, (prev, curr) in enumerate(zip(previous_states, current_states)):
                        if curr and not prev:  # Button was just pressed
                            button_number = i + 1
                            self.trigger_key = button_number
                            self.config['Settings']['trigger_key'] = str(button_number)
                            with open('config.ini', 'w') as f:
                                self.config.write(f)
                            
                            self.root.after(0, lambda b=button_number: self.status_label.config(
                                text=f"Current trigger: Mouse Button {b}"))
                            self.root.after(0, lambda: self.config_button.config(
                                text="Press to configure new button"))
                            
                            self.capturing = False
                            return
                    
                    previous_states = current_states
                    time.sleep(0.01)
                except Exception as e:
                    print(f"Error in capture loop: {e}")
                    time.sleep(0.1)
                
        except Exception as e:
            print(f"Error in capture_mouse_button: {e}")
            self.capturing = False
    
    def perform_key_sequence(self):
        """Perform the space+shift key sequence"""
        try:
            # Release any potentially stuck keys first
            keyboard.release('space')
            keyboard.release('shift')
            
            # Calculate adjusted delay based on system speed
            base_delay = self.key_delay / self.speeds[self.speed_var.get()]
            adjusted_delay = base_delay * self.system_delay_compensation
            
            # Ensure minimum delay
            adjusted_delay = max(adjusted_delay, 0.001)
            
            print(f"Using adjusted delay: {adjusted_delay}")
            
            # Complete sequence with proper delays
            keyboard.press('space')
            print("Space pressed")
            time.sleep(adjusted_delay)
            
            keyboard.press('shift')
            print("Shift pressed")
            time.sleep(adjusted_delay)
            
            keyboard.release('space')
            print("Space released")
            time.sleep(adjusted_delay / 2)
            
            keyboard.release('shift')
            print("Shift released")
            time.sleep(adjusted_delay / 2)
            
            print("Sequence complete!")
            
        except Exception as e:
            print(f"Error in key sequence: {e}")
            # Emergency key release
            try:
                keyboard.release('space')
                keyboard.release('shift')
            except:
                pass
    
    def monitor_mouse(self):
        """Monitor for mouse button presses"""
        try:
            print(f"Monitoring started - System: {sys.platform}, Windows version: {sys.getwindowsversion()}")
            print(f"Using trigger key: {self.trigger_key}")
            last_press_time = 0
            button_was_pressed = False
            
            while self.running:
                try:
                    current_state = win32api.GetAsyncKeyState(self.trigger_key)
                    current_time = time.perf_counter()  # More precise timing
                    
                    # Check if button is pressed (negative state indicates pressed)
                    if current_state < 0:  
                        if not button_was_pressed or (
                            self.repeat_enabled and 
                            current_time - last_press_time >= (self.key_delay * self.system_delay_compensation)
                        ):
                            print(f"Button {self.trigger_key} pressed at {current_time}")
                            self.perform_key_sequence()
                            last_press_time = current_time
                        button_was_pressed = True
                    else:
                        button_was_pressed = False
                    
                    # Dynamic sleep based on system performance
                    time.sleep(min(0.01 * self.system_delay_compensation, 0.01))
                    
                except Exception as e:
                    print(f"Error in monitor_mouse loop: {e}")
                    time.sleep(0.1)
                    
        except Exception as e:
            print(f"Fatal error in monitor_mouse: {e}")
        finally:
            self.running = False
            self.root.after(0, lambda: self.toggle_button.config(text="Start"))
            try:
                keyboard.release('space')
                keyboard.release('shift')
            except:
                pass
    
    def toggle_script(self):
        try:
            if not self.running:
                print("Starting script...")
                self.running = True
                self.toggle_button.config(text="Stop")
                self.monitor_thread = threading.Thread(target=self.monitor_mouse)
                self.monitor_thread.daemon = True
                self.monitor_thread.start()
                print(f"Script started - using button {self.trigger_key}")
            else:
                print("Stopping script...")
                self.running = False
                self.toggle_button.config(text="Start")
                self.monitor_thread = None
                print("Script stopped")
        except Exception as e:
            print(f"Error toggling script: {e}")
    
    def save_settings(self):
        try:
            self.repeat_enabled = self.repeat_var.get()
            self.key_delay = self.delay_var.get()
            self.speed_multiplier = self.speeds[self.speed_var.get()]  # Get multiplier from selected speed
            self.config['Settings']['repeat_enabled'] = str(self.repeat_enabled).lower()
            self.config['Settings']['key_delay'] = str(self.key_delay)
            self.config['Settings']['speed_multiplier'] = str(self.speed_multiplier)
            with open('config.ini', 'w') as f:
                self.config.write(f)
        except Exception as e:
            print(f"Error saving settings: {e}")

def main():
    try:
        root = tk.Tk()
        app = KeyPresserApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Fatal error in main: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
