import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pyautogui
import pyperclip
import time
import threading
import json
from pynput import mouse, keyboard
import queue
from PIL import ImageGrab
import re
from collections import deque

class AdvancedRecorderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AutomatePro - Universal Smart Recorder")
        self.root.geometry("950x950")
        
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.05
        
        self.recording = False
        self.playing = False
        self.recorded_actions = []
        self.action_queue = queue.Queue()
        
        self.action_lock = threading.Lock()
        
        self.mouse_pressed = False
        self.selection_start = None
        self.selection_end = None
        self.last_scroll_time = 0
        self.is_dragging = False
        self.drag_start_time = None
        self.scroll_before_selection = []
        
        self.current_combo = set()
        self.ctrl_pressed = False
        self.shift_pressed = False
        self.alt_pressed = False
        self.last_key_time = 0
        self.combo_detected = False
        self.pending_combo_action = None
        
        self.last_clipboard = ""
        self.clipboard_before_copy = ""
        self.clipboard_history = deque(maxlen=10)
        self.selection_context = {}
        
        self.mouse_listener = None
        self.keyboard_listener = None
        
        self.context_actions = []
        self.last_action_time = 0
        self.copy_action_id = 0
        
        self.max_actions = 10000
        
        self.debug_mode = False
        self.error_log = []
        

        self.current_context = "absolute"  # Default to absolute for safety
        self.current_app_type = "unknown"
        self.last_arrow_key_time = 0
        self.arrow_key_count = 0
        self.last_copy_paste_time = 0
        self.context_switch_times = deque(maxlen=5)
        self.recent_actions = deque(maxlen=20)
        self.last_window_check = 0
        self.cached_window_type = "unknown"
        

        self.clipboard_snapshots = {}  # Store clipboard content by action ID
        
        self.setup_gui()
        
    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="AutomatePro - Universal Smart Recorder", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=4, pady=10)
        

        context_frame = ttk.Frame(main_frame)
        context_frame.grid(row=1, column=0, columnspan=4, pady=5)
        ttk.Label(context_frame, text="Mode:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        self.context_label = ttk.Label(context_frame, text="ABSOLUTE", 
                                       font=('Arial', 10), foreground='orange')
        self.context_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(context_frame, text="| App:", font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        self.app_type_label = ttk.Label(context_frame, text="UNKNOWN", 
                                        font=('Arial', 10), foreground='gray')
        self.app_type_label.pack(side=tk.LEFT, padx=5)
        
        control_frame = ttk.LabelFrame(main_frame, text="Controls", padding="10")
        control_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=10)
        
        self.record_btn = ttk.Button(control_frame, text="Start Recording", 
                                     command=self.toggle_recording, width=20)
        self.record_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.play_btn = ttk.Button(control_frame, text="Play Once", 
                                   command=self.play_sequence, width=20)
        self.play_btn.grid(row=0, column=1, padx=5, pady=5)
        
        self.loop_play_btn = ttk.Button(control_frame, text="Loop Play", 
                                        command=self.start_loop_play, width=20)
        self.loop_play_btn.grid(row=0, column=2, padx=5, pady=5)
        
        self.clear_btn = ttk.Button(control_frame, text="Clear", 
                                    command=self.clear_sequence, width=20)
        self.clear_btn.grid(row=0, column=3, padx=5, pady=5)
        
        loop_frame = ttk.Frame(control_frame)
        loop_frame.grid(row=1, column=0, columnspan=4, pady=(10,0))
        ttk.Label(loop_frame, text="Repeat:").pack(side=tk.LEFT)
        self.loop_count_var = tk.IntVar(value=3)
        ttk.Spinbox(loop_frame, from_=1, to=100, width=5, textvariable=self.loop_count_var).pack(side=tk.LEFT, padx=5)
        ttk.Label(loop_frame, text="times").pack(side=tk.LEFT)
        
        options_frame = ttk.LabelFrame(main_frame, text="Recording Options", padding="10")
        options_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        self.record_keyboard_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Keyboard", 
                       variable=self.record_keyboard_var).grid(row=0, column=0, padx=5)
        
        self.record_mouse_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Mouse", 
                       variable=self.record_mouse_var).grid(row=0, column=1, padx=5)
        
        self.record_scroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Scroll", 
                       variable=self.record_scroll_var).grid(row=0, column=2, padx=5)
        
        self.record_clipboard_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Clipboard", 
                       variable=self.record_clipboard_var).grid(row=0, column=3, padx=5)
        

        mode_frame = ttk.LabelFrame(main_frame, text="Recording Mode", padding="10")
        mode_frame.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        self.mode_var = tk.StringVar(value="smart")
        
        ttk.Radiobutton(mode_frame, text="🧠 Smart Auto-Detect (Recommended)", 
                       variable=self.mode_var, value="smart",
                       command=self.on_mode_change).grid(row=0, column=0, sticky=tk.W, padx=10)
        
        ttk.Radiobutton(mode_frame, text="📊 Desktop Excel Mode (Relative Navigation)", 
                       variable=self.mode_var, value="excel",
                       command=self.on_mode_change).grid(row=1, column=0, sticky=tk.W, padx=10)
        
        ttk.Radiobutton(mode_frame, text="🎯 Absolute Mode (Web Apps, Click-based)", 
                       variable=self.mode_var, value="absolute",
                       command=self.on_mode_change).grid(row=2, column=0, sticky=tk.W, padx=10)
        

        smart_frame = ttk.Frame(mode_frame)
        smart_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), padx=20, pady=5)
        
        self.web_app_priority_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(smart_frame, 
                       text="🌐 Always use Absolute mode for browsers (Google Sheets, Apollo, etc.)",
                       variable=self.web_app_priority_var).pack(anchor=tk.W)
        
        self.excel_click_record_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(smart_frame, 
                       text="📌 Record clicks even in Excel mode (for menus/ribbons)",
                       variable=self.excel_click_record_var).pack(anchor=tk.W)
        
        file_frame = ttk.LabelFrame(main_frame, text="File Operations", padding="10")
        file_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        self.save_btn = ttk.Button(file_frame, text="Save", 
                                   command=self.save_sequence, width=20)
        self.save_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.load_btn = ttk.Button(file_frame, text="Load", 
                                   command=self.load_sequence, width=20)
        self.load_btn.grid(row=0, column=1, padx=5, pady=5)
        
        speed_label = ttk.Label(file_frame, text="Playback Speed:")
        speed_label.grid(row=0, column=2, padx=5)
        
        self.speed_var = tk.DoubleVar(value=1.0)
        self.speed_scale = ttk.Scale(file_frame, from_=0.5, to=3.0, 
                                     variable=self.speed_var, orient=tk.HORIZONTAL, 
                                     length=150)
        self.speed_scale.grid(row=0, column=3, padx=5)
        
        self.speed_label = ttk.Label(file_frame, text="1.0x")
        self.speed_label.grid(row=0, column=4, padx=5)
        
        def update_speed_label(value):
            self.speed_label.config(text=f"{float(value):.1f}x")
        
        self.speed_scale.config(command=update_speed_label)
        
        sequence_frame = ttk.LabelFrame(main_frame, text="Recorded Actions", padding="10")
        sequence_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.sequence_text = scrolledtext.ScrolledText(sequence_frame, width=80, height=12)
        self.sequence_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        

        debug_frame = ttk.LabelFrame(main_frame, text="Activity Log", padding="10")
        debug_frame.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        self.debug_text = scrolledtext.ScrolledText(debug_frame, width=80, height=5)
        self.debug_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.status_var = tk.StringVar(value="Ready - Smart Mode Enabled")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=8, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        instructions = ttk.Label(main_frame, 
                                text="Press F9 to stop recording/playback • Developed by Nader Mahbub Khan for Chandan Chakraborty",
                                font=('Arial', 9), foreground='gray')
        instructions.grid(row=9, column=0, columnspan=4, pady=5)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        sequence_frame.columnconfigure(0, weight=1)
        sequence_frame.rowconfigure(0, weight=1)
        debug_frame.columnconfigure(0, weight=1)
        
        self.setup_global_hotkey()
    
    def on_mode_change(self):
        mode = self.mode_var.get()
        if mode == "smart":
            self.status_var.set("Smart Mode: Auto-detecting context")
            self.log_debug("🧠 Smart mode enabled")
        elif mode == "excel":
            self.status_var.set("Excel Mode: Recording keyboard navigation, minimal clicks")
            self.current_context = "excel"
            self.log_debug("📊 Excel mode enabled")
        elif mode == "absolute":
            self.status_var.set("Absolute Mode: Recording all mouse positions and clicks")
            self.current_context = "absolute"
            self.log_debug("🎯 Absolute mode enabled")
    
    def detect_window_type(self):
        
        current_time = time.time()
        

        if current_time - self.last_window_check < 2.0:
            return self.cached_window_type
        
        self.last_window_check = current_time
        
        try:

            try:
                import pygetwindow as gw
                active_window = gw.getActiveWindow()
                if active_window:
                    title = active_window.title.lower()
                    

                    browsers = ['chrome', 'firefox', 'edge', 'safari', 'brave', 'opera', 'vivaldi']
                    if any(browser in title for browser in browsers):

                        if 'sheets' in title or 'google sheets' in title:
                            self.cached_window_type = "web_sheets"
                            return "web_sheets"
                        elif 'apollo' in title:
                            self.cached_window_type = "web_crm"
                            return "web_crm"
                        elif 'airtable' in title:
                            self.cached_window_type = "web_database"
                            return "web_database"
                        else:
                            self.cached_window_type = "web_browser"
                            return "web_browser"
                    

                    if 'excel' in title and 'microsoft' in title:
                        self.cached_window_type = "desktop_excel"
                        return "desktop_excel"
                    

                    self.cached_window_type = "desktop_app"
                    return "desktop_app"
            except ImportError:

                pass
        except Exception as e:
            self.log_error("Window detection", e)
        
        self.cached_window_type = "unknown"
        return "unknown"
    
    def detect_context(self, action_type, key=None):
        
        

        if self.mode_var.get() == "excel":
            return "excel"
        elif self.mode_var.get() == "absolute":
            return "absolute"
        

        window_type = self.detect_window_type()
        

        self.root.after(0, lambda: self.app_type_label.config(
            text=window_type.replace('_', ' ').upper()
        ))
        

        if self.web_app_priority_var.get():
            if window_type in ['web_browser', 'web_sheets', 'web_crm', 'web_database']:
                self.current_context = "absolute"
                self.root.after(0, lambda: self.context_label.config(
                    text="ABSOLUTE (WEB)", foreground='orange'
                ))
                return "absolute"
        

        if window_type == "desktop_excel":
            current_time = time.time()
            

            self.recent_actions.append({
                'type': action_type,
                'key': key,
                'time': current_time
            })
            

            recent_arrows = sum(1 for a in self.recent_actions 
                               if a['time'] > current_time - 5 
                               and a.get('key') in ['up', 'down', 'left', 'right'])
            

            recent_copy_paste = sum(1 for a in self.recent_actions 
                                   if a['time'] > current_time - 3 
                                   and a['type'] in ['copy', 'paste'])
            

            if recent_arrows >= 2 or (recent_copy_paste > 0 and recent_arrows > 0):
                self.current_context = "excel"
                self.root.after(0, lambda: self.context_label.config(
                    text="EXCEL", foreground='green'
                ))
                return "excel"
        

        self.current_context = "absolute"
        self.root.after(0, lambda: self.context_label.config(
            text="ABSOLUTE", foreground='orange'
        ))
        return "absolute"
    
    def log_debug(self, message):
        
        timestamp = time.strftime('%H:%M:%S')
        self.debug_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.debug_text.see(tk.END)

        lines = self.debug_text.get('1.0', tk.END).split('\n')
        if len(lines) > 100:
            self.debug_text.delete('1.0', f'{len(lines)-100}.0')
    
    def log_error(self, context, error):
        error_msg = f"[{time.strftime('%H:%M:%S')}] {context}: {str(error)}"
        self.error_log.append(error_msg)
        self.log_debug(f"❌ ERROR: {error_msg}")
        if len(self.error_log) > 100:
            self.error_log = self.error_log[-100:]
    
    def add_action(self, action):
        with self.action_lock:
            self.recorded_actions.append(action)
            if len(self.recorded_actions) > self.max_actions:
                self.recorded_actions = self.recorded_actions[-self.max_actions:]
                self.root.after(0, lambda: messagebox.showwarning(
                    "Warning", 
                    f"Action limit reached. Keeping last {self.max_actions} actions."
                ))
        self.root.after(0, self.update_sequence_display)
    
    def safe_update_status(self, message):
        try:
            self.root.after(0, lambda: self.status_var.set(message))
        except Exception as e:
            self.log_error("Status update", e)
    
    def setup_global_hotkey(self):
        def on_press(key):
            try:
                if key == keyboard.Key.f9:
                    if self.recording:
                        self.root.after(0, self.stop_recording)
                    if self.playing:
                        self.playing = False
            except:
                pass
                
        listener = keyboard.Listener(on_press=on_press)
        listener.start()
        
    def get_clipboard_safe(self, max_retries=3, delay=0.15):
        
        for attempt in range(max_retries):
            try:
                content = pyperclip.paste()
                return content if content else ""
            except Exception as e:
                if attempt == max_retries - 1:
                    self.log_error("Clipboard access", e)
                    return ""
                time.sleep(delay)
        return ""
    
    def set_clipboard_safe(self, content, max_retries=3, delay=0.15):
        
        for attempt in range(max_retries):
            try:
                pyperclip.copy(content)
                time.sleep(0.1)  # Give system time to update
                return True
            except Exception as e:
                if attempt == max_retries - 1:
                    self.log_error("Clipboard set", e)
                    return False
                time.sleep(delay)
        return False
    
    def monitor_clipboard(self):
        
        while self.recording:
            try:
                current_clipboard = self.get_clipboard_safe()
                if current_clipboard and current_clipboard != self.last_clipboard:
                    if self.record_clipboard_var.get():
                        self.last_clipboard = current_clipboard
                        self.clipboard_history.append(current_clipboard)
            except Exception as e:
                self.log_error("Clipboard monitor", e)
            time.sleep(0.25)
            
    def toggle_recording(self):
        if not self.recording:
            self.start_recording()
        else:
            self.stop_recording()
    
    def start_recording(self):
        self.recording = True
        self.record_btn.config(text="⏹ Stop Recording")
        self.play_btn.config(state='disabled')
        self.loop_play_btn.config(state='disabled')
        self.clear_btn.config(state='disabled')
        self.save_btn.config(state='disabled')
        self.load_btn.config(state='disabled')
        
        mode_text = self.mode_var.get().upper()
        self.status_var.set(f"🔴 Recording ({mode_text} mode) - Press F9 to stop")
        
        with self.action_lock:
            self.recorded_actions = []
        self.update_sequence_display()
        
        self.mouse_pressed = False
        self.selection_start = None
        self.selection_end = None
        self.is_dragging = False
        self.drag_start_time = None
        self.context_actions = []
        self.current_combo = set()
        self.ctrl_pressed = False
        self.shift_pressed = False
        self.alt_pressed = False
        self.combo_detected = False
        self.selection_context = {}
        self.clipboard_history.clear()
        self.clipboard_snapshots.clear()
        self.pending_combo_action = None
        self.copy_action_id = 0
        self.last_arrow_key_time = 0
        self.arrow_key_count = 0
        self.recent_actions.clear()
        
        self.last_clipboard = self.get_clipboard_safe()
        self.clipboard_before_copy = self.last_clipboard
        self.log_debug("Developed by Nader Mahbub Khan")
        self.log_debug("Phone: 01642817116")
        self.log_debug("==============================")
        self.log_debug("🎬 Recording started")
        
        clipboard_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
        clipboard_thread.start()
        
        def on_click(x, y, button, pressed):
            if not self.recording or not self.record_mouse_var.get():
                return
            
            current_time = time.time()
            
            if pressed:
                self.mouse_pressed = True
                self.selection_start = (x, y, current_time)
                self.drag_start_time = current_time
                

                context = self.detect_context('click')
                

                if button in [mouse.Button.right, mouse.Button.middle]:
                    action = {
                        'type': 'right_click' if button == mouse.Button.right else 'middle_click',
                        'x': x,
                        'y': y,
                        'time': current_time,
                        'context': context
                    }
                    self.add_action(action)
                    
            else:
                if self.mouse_pressed and self.selection_start:
                    self.selection_end = (x, y, current_time)
                    
                    start_x, start_y, start_time = self.selection_start
                    distance = ((x - start_x)**2 + (y - start_y)**2)**0.5
                    duration = current_time - self.drag_start_time
                    
                    context = self.detect_context('click')
                    

                    if distance > 10 and duration > 0.1:

                        if context != "excel" or self.excel_click_record_var.get():
                            action = {
                                'type': 'selection',
                                'start_x': start_x,
                                'start_y': start_y,
                                'end_x': x,
                                'end_y': y,
                                'duration': duration,
                                'time': current_time,
                                'context': context
                            }
                            self.add_action(action)
                    else:

                        if context != "excel" or self.excel_click_record_var.get():
                            action = {
                                'type': 'click',
                                'x': start_x,
                                'y': start_y,
                                'button': str(button),
                                'time': current_time,
                                'context': context
                            }
                            self.add_action(action)
                
                self.mouse_pressed = False
                self.selection_start = None
                self.is_dragging = False
                
        def on_scroll(x, y, dx, dy):
            if not self.recording or not self.record_scroll_var.get():
                return
            
            current_time = time.time()
            context = self.detect_context('scroll')
            

            action = {
                'type': 'scroll',
                'x': x,
                'y': y,
                'dx': dx,
                'dy': dy,
                'time': current_time,
                'context': context
            }
            self.add_action(action)
            
        def on_key_press(key):
            if not self.recording or not self.record_keyboard_var.get():
                return
            if key == keyboard.Key.f9:
                return

            vk = getattr(key, 'vk', None)
            current_time = time.time()

            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, keyboard.Key.ctrl]:
                self.ctrl_pressed = True
                return
            elif key in [keyboard.Key.shift_l, keyboard.Key.shift_r, keyboard.Key.shift]:
                self.shift_pressed = True
                return
            elif key in [keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt, keyboard.Key.alt_gr]:
                self.alt_pressed = True
                return


            key_name = None
            if hasattr(key, 'name'):
                key_name = key.name.lower()
            
            if key_name in ['up', 'down', 'left', 'right']:
                self.last_arrow_key_time = current_time
                self.arrow_key_count += 1

            if self.ctrl_pressed and vk is not None:
                if vk == 67:  # 'C' - Copy
                    self.clipboard_before_copy = self.get_clipboard_safe()
                    self.copy_action_id += 1
                    current_id = self.copy_action_id
                    self.last_copy_paste_time = current_time
                    
                    context = self.detect_context('copy')
                    
                    action = {
                        'type': 'copy',
                        'time': current_time,
                        '_id': current_id,
                        'context': context
                    }
                    self.add_action(action)


                    def delayed_capture(copy_id):
                        time.sleep(1.5)  # Increased delay for reliability
                        new_clip = self.get_clipboard_safe()
                        if new_clip and new_clip != self.clipboard_before_copy:

                            self.clipboard_snapshots[copy_id] = new_clip
                            
                            with self.action_lock:
                                for act in reversed(self.recorded_actions):
                                    if act.get('_id') == copy_id:
                                        preview = new_clip[:50] + "..." if len(new_clip) > 50 else new_clip
                                        act['preview'] = preview
                                        act['content'] = new_clip  # Store full content
                                        break
                            self.root.after(0, self.update_sequence_display)

                    threading.Thread(target=delayed_capture, args=(current_id,), daemon=True).start()
                    return

                elif vk == 86:  # 'V' - Paste
                    self.last_copy_paste_time = current_time
                    context = self.detect_context('paste')
                    

                    current_clip = self.get_clipboard_safe()
                    
                    action = {
                        'type': 'paste',
                        'time': current_time,
                        'context': context,
                        'content': current_clip,  # Store what will be pasted
                        'dynamic': True
                    }
                    
                    self.add_action(action)
                    self.log_debug(f"📋 Paste recorded (will use current clipboard)")
                    return

                elif vk == 88:  # 'X' - Cut
                    self.clipboard_before_copy = self.get_clipboard_safe()
                    context = self.detect_context('cut')
                    cut_id = self.copy_action_id + 1
                    self.copy_action_id = cut_id
                    
                    action = {
                        'type': 'cut',
                        'time': current_time,
                        'context': context,
                        '_id': cut_id
                    }
                    self.add_action(action)

                    def capture_cut(cut_id):
                        time.sleep(1.2)
                        new_clip = self.get_clipboard_safe()
                        if new_clip and new_clip != self.clipboard_before_copy:
                            self.clipboard_snapshots[cut_id] = new_clip
                            with self.action_lock:
                                for act in reversed(self.recorded_actions):
                                    if act.get('_id') == cut_id:
                                        preview = new_clip[:50] + "..." if len(new_clip) > 50 else new_clip
                                        act['preview'] = preview
                                        act['content'] = new_clip
                                        break
                            self.root.after(0, self.update_sequence_display)

                    threading.Thread(target=capture_cut, args=(cut_id,), daemon=True).start()
                    return
                

                elif vk == 78:  # 'N' - New
                    context = self.detect_context('key', 'n')
                    action = {'type': 'new_file', 'time': current_time, 'context': context}
                    self.add_action(action)
                    return
                elif vk == 84:  # 'T' - New tab
                    context = self.detect_context('key', 't')
                    action = {'type': 'new_tab', 'time': current_time, 'context': context}
                    self.add_action(action)
                    return
                elif vk == 65:  # 'A' - Select all
                    action = {'type': 'select_all', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 90:  # 'Z' - Undo
                    action = {'type': 'undo', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 89:  # 'Y' - Redo
                    action = {'type': 'redo', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 83:  # 'S' - Save
                    action = {'type': 'save', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 70:  # 'F' - Find
                    action = {'type': 'find', 'time': current_time}
                    self.add_action(action)
                    return


            key_repr = None

            if hasattr(key, 'name'):
                name = key.name.lower()
                special_keys = {
                    'space': 'space',
                    'enter': 'enter',
                    'return': 'enter',
                    'tab': 'tab',
                    'backspace': 'backspace',
                    'delete': 'delete',
                    'home': 'home',
                    'end': 'end',
                    'page_up': 'pageup',
                    'page_down': 'pagedown',
                    'up': 'up',
                    'down': 'down',
                    'left': 'left',
                    'right': 'right',
                    'escape': 'esc',
                    'esc': 'esc',
                    'insert': 'insert',
                    'menu': 'menu',
                    'caps_lock': 'capslock',
                    'num_lock': 'numlock',
                    'scroll_lock': 'scrolllock',
                    'print_screen': 'printscreen',
                    'pause': 'pause',
                    'f1': 'f1', 'f2': 'f2', 'f3': 'f3', 'f4': 'f4',
                    'f5': 'f5', 'f6': 'f6', 'f7': 'f7', 'f8': 'f8',
                    'f9': 'f9', 'f10': 'f10', 'f11': 'f11', 'f12': 'f12',
                }
                if name in special_keys:
                    key_repr = special_keys[name]
                else:
                    return

            elif hasattr(key, 'char') and key.char and (key.char.isprintable() or key.char in ['\t', '\n', '\r']):
                key_repr = key.char

            elif vk is not None:
                if 65 <= vk <= 90:
                    key_repr = chr(vk).lower()
                elif 48 <= vk <= 57:
                    key_repr = chr(vk)

            if key_repr is not None:
                context = self.detect_context('key', key_repr)
                
                action = {
                    'type': 'key',
                    'key': key_repr,
                    'ctrl': self.ctrl_pressed,
                    'shift': self.shift_pressed,
                    'alt': self.alt_pressed,
                    'time': current_time,
                    'context': context
                }
                self.add_action(action)

        def on_key_release(key):
            if not self.recording:
                return
            
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, keyboard.Key.ctrl]:
                self.ctrl_pressed = False
            elif key in [keyboard.Key.shift, keyboard.Key.shift_r, keyboard.Key.shift_l]:
                self.shift_pressed = False
            elif key in [keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt, keyboard.Key.alt_gr]:
                self.alt_pressed = False
                
        self.mouse_listener = mouse.Listener(
            on_click=on_click,
            on_scroll=on_scroll
        )
        self.keyboard_listener = keyboard.Listener(
            on_press=on_key_press,
            on_release=on_key_release
        )
        
        self.mouse_listener.start()
        self.keyboard_listener.start()
        
    def stop_recording(self):
        self.recording = False
        self.record_btn.config(text="🔴 Start Recording")
        self.play_btn.config(state='normal')
        self.loop_play_btn.config(state='normal')
        self.clear_btn.config(state='normal')
        self.save_btn.config(state='normal')
        self.load_btn.config(state='normal')
        
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener.join()
        if self.keyboard_listener:
            self.keyboard_listener.stop()
            self.keyboard_listener.join()
        
        with self.action_lock:
            action_count = len(self.recorded_actions)
        
        self.log_debug(f"⏹ Recording stopped - {action_count} actions captured")
        self.status_var.set(f"Recording stopped. {action_count} actions recorded.")
    
    def play_sequence(self, loop_count=1):
        with self.action_lock:
            if not self.recorded_actions:
                messagebox.showwarning("Warning", "No actions recorded")
                return
            actions_to_play = self.recorded_actions.copy()
            
        def play_thread():
            original_failsafe = pyautogui.FAILSAFE
            pyautogui.FAILSAFE = False
            try:
                self.playing = True
                self.play_btn.config(state='disabled')
                self.loop_play_btn.config(state='disabled')
                self.record_btn.config(state='disabled')
                self.clear_btn.config(state='disabled')
                self.save_btn.config(state='disabled')
                self.load_btn.config(state='disabled')
                
                self.safe_update_status(f"▶ Playing sequence... (Press F9 to stop)")
                time.sleep(2)
                
                for loop in range(loop_count):
                    if not self.playing:
                        break
                    if loop_count > 1:
                        self.safe_update_status(f"🔄 Loop {loop+1}/{loop_count}...")
                        time.sleep(0.5)
                    
                    speed = self.speed_var.get()
                    prev_time = None
                    screen_width, screen_height = pyautogui.size()
                    
                    special_pyautogui_keys = {
                        'space', 'enter', 'tab', 'backspace', 'delete', 'home', 'end',
                        'pageup', 'pagedown', 'up', 'down', 'left', 'right', 'esc',
                        'insert', 'menu', 'capslock', 'numlock', 'scrolllock', 'printscreen',
                        'pause', 'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12'
                    }
                    
                    for i, action in enumerate(actions_to_play):
                        if not self.playing:
                            break
                            
                        if prev_time:
                            delay = (action['time'] - prev_time) / speed
                            delay = min(delay, 10)
                            if delay > 0.01:
                                time.sleep(delay)
                            
                        prev_time = action['time']
                        
                        try:
                            action_type = action['type']
                            context = action.get('context', 'absolute')
                            
                            def safe_coords(x, y):
                                x = max(0, min(screen_width - 1, x))
                                y = max(0, min(screen_height - 1, y))
                                return int(x), int(y)
                            

                            if action_type == 'click':
                                x, y = safe_coords(action['x'], action['y'])
                                pyautogui.moveTo(x, y, duration=0.2, tween=pyautogui.easeInOutQuad)
                                pyautogui.click()
                                self.safe_update_status(f"Playing: Click at ({x}, {y})")
                                
                            elif action_type == 'right_click':
                                x, y = safe_coords(action['x'], action['y'])
                                pyautogui.moveTo(x, y, duration=0.2, tween=pyautogui.easeInOutQuad)
                                pyautogui.rightClick()
                                self.safe_update_status(f"Playing: Right-click at ({x}, {y})")
                                
                            elif action_type == 'middle_click':
                                x, y = safe_coords(action['x'], action['y'])
                                pyautogui.moveTo(x, y, duration=0.2, tween=pyautogui.easeInOutQuad)
                                pyautogui.middleClick()
                                self.safe_update_status(f"Playing: Middle-click at ({x}, {y})")
                                
                            elif action_type == 'selection':
                                start_x, start_y = safe_coords(action['start_x'], action['start_y'])
                                end_x, end_y = safe_coords(action['end_x'], action['end_y'])
                                pyautogui.moveTo(start_x, start_y, duration=0.2, tween=pyautogui.easeInOutQuad)
                                pyautogui.mouseDown()
                                time.sleep(0.05)
                                pyautogui.moveTo(end_x, end_y, duration=0.3, tween=pyautogui.easeInOutQuad)
                                time.sleep(0.05)
                                pyautogui.mouseUp()
                                self.safe_update_status("Playing: Selection/Drag")
                                
                            elif action_type == 'scroll':
                                x, y = safe_coords(action['x'], action['y'])
                                pyautogui.moveTo(x, y, duration=0.1)
                                scroll_amount = int(action['dy'] * 120)
                                pyautogui.scroll(scroll_amount)
                                self.safe_update_status("Playing: Scroll")
                                time.sleep(0.05)
                                
                            elif action_type == 'copy':
                                pyautogui.hotkey('ctrl', 'c')
                                self.safe_update_status("Playing: Copy (Ctrl+C)")
                                time.sleep(0.8)  # Wait for clipboard to update
                                
                            elif action_type == 'paste':

                                current_clip = self.get_clipboard_safe()
                                


                                if 'content' in action and action['content']:
                                    if not action.get('dynamic', False):

                                        self.set_clipboard_safe(action['content'])
                                        time.sleep(0.2)
                                
                                preview = current_clip[:30] + "..." if len(current_clip) > 30 else current_clip
                                self.safe_update_status(f"Playing: Paste [{preview}]")
                                pyautogui.hotkey('ctrl', 'v')
                                time.sleep(0.4)
                                
                            elif action_type == 'cut':
                                pyautogui.hotkey('ctrl', 'x')
                                self.safe_update_status("Playing: Cut (Ctrl+X)")
                                time.sleep(0.6)
                            
                            elif action_type == 'new_file':
                                pyautogui.hotkey('ctrl', 'n')
                                self.safe_update_status("Playing: New File (Ctrl+N)")
                                time.sleep(0.5)
                            
                            elif action_type == 'new_tab':
                                pyautogui.hotkey('ctrl', 't')
                                self.safe_update_status("Playing: New Tab (Ctrl+T)")
                                time.sleep(0.5)
                                
                            elif action_type == 'select_all':
                                pyautogui.hotkey('ctrl', 'a')
                                self.safe_update_status("Playing: Select All (Ctrl+A)")
                                time.sleep(0.1)
                                
                            elif action_type == 'undo':
                                pyautogui.hotkey('ctrl', 'z')
                                self.safe_update_status("Playing: Undo (Ctrl+Z)")
                                time.sleep(0.1)
                                
                            elif action_type == 'redo':
                                pyautogui.hotkey('ctrl', 'y')
                                self.safe_update_status("Playing: Redo (Ctrl+Y)")
                                time.sleep(0.1)
                                
                            elif action_type == 'save':
                                pyautogui.hotkey('ctrl', 's')
                                self.safe_update_status("Playing: Save (Ctrl+S)")
                                time.sleep(0.1)
                                
                            elif action_type == 'find':
                                pyautogui.hotkey('ctrl', 'f')
                                self.safe_update_status("Playing: Find (Ctrl+F)")
                                time.sleep(0.1)
                                
                            elif action_type == 'key':
                                key = action['key']
                                modifiers = []
                                if action.get('ctrl'):
                                    modifiers.append('ctrl')
                                if action.get('shift'):
                                    modifiers.append('shift')
                                if action.get('alt'):
                                    modifiers.append('alt')
                                    
                                if modifiers:
                                    pyautogui.hotkey(*modifiers, key)
                                elif key in special_pyautogui_keys:
                                    pyautogui.press(key)
                                else:
                                    pyautogui.write(key)
                                
                                ctx_tag = f" [{context.upper()}]" if key in ['up','down','left','right'] else ""
                                self.safe_update_status(f"Playing: Key '{key}'{ctx_tag}")
                                

                                if key in ['up', 'down', 'left', 'right'] and context == 'excel':
                                    time.sleep(0.3)
                                
                        except Exception as e:
                            self.log_error(f"Playing action {i}", e)
                            
                self.playing = False
                self.play_btn.config(state='normal')
                self.loop_play_btn.config(state='normal')
                self.record_btn.config(state='normal')
                self.clear_btn.config(state='normal')
                self.save_btn.config(state='normal')
                self.load_btn.config(state='normal')
                self.safe_update_status("✅ Playback completed")
                
            finally:
                pyautogui.FAILSAFE = original_failsafe
                
        threading.Thread(target=play_thread, daemon=True).start()
        
    def start_loop_play(self):
        try:
            count = self.loop_count_var.get()
            if count < 1:
                count = 1
            self.play_sequence(loop_count=count)
        except Exception as e:
            messagebox.showerror("Error", f"Invalid loop count: {e}")
        
    def clear_sequence(self):
        with self.action_lock:
            self.recorded_actions = []
        self.sequence_text.delete('1.0', tk.END)
        self.debug_text.delete('1.0', tk.END)
        self.clipboard_snapshots.clear()
        self.current_context = "absolute"
        self.context_label.config(text="ABSOLUTE", foreground="orange")
        self.status_var.set("Sequence cleared")
        
    def save_sequence(self):
        with self.action_lock:
            if not self.recorded_actions:
                messagebox.showwarning("Warning", "No actions to save")
                return
            actions_to_save = self.recorded_actions.copy()
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(actions_to_save, f, indent=2, ensure_ascii=False)
                self.status_var.set(f"Sequence saved to {filename}")
                messagebox.showinfo("Success", "Sequence saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {str(e)}")
                
    def load_sequence(self):
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    loaded_actions = json.load(f)
                with self.action_lock:
                    self.recorded_actions = loaded_actions
                self.update_sequence_display()
                self.status_var.set(f"Sequence loaded from {filename}")
                messagebox.showinfo("Success", f"Loaded {len(loaded_actions)} actions")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load: {str(e)}")
                
    def update_sequence_display(self):
        self.sequence_text.delete('1.0', tk.END)
        with self.action_lock:
            actions_copy = self.recorded_actions.copy()
        
        for i, action in enumerate(actions_copy, 1):
            action_type = action.get('type', 'unknown')
            context = action.get('context', 'absolute')
            

            if context == 'excel':
                ctx_icon = "📊"
            elif context == 'absolute':
                ctx_icon = "🎯"
            else:
                ctx_icon = "🔀"
            
            if action_type == 'click':
                text = f"{i}. {ctx_icon} Click at ({action['x']}, {action['y']})\n"
            elif action_type == 'right_click':
                text = f"{i}. {ctx_icon} Right-click at ({action['x']}, {action['y']})\n"
            elif action_type == 'middle_click':
                text = f"{i}. {ctx_icon} Middle-click at ({action['x']}, {action['y']})\n"
            elif action_type == 'selection':
                text = f"{i}. {ctx_icon} Selection from ({action['start_x']}, {action['start_y']}) to ({action['end_x']}, {action['end_y']})\n"
            elif action_type == 'scroll':
                direction = "down" if action['dy'] < 0 else "up"
                text = f"{i}. {ctx_icon} Scroll {direction}\n"
            elif action_type == 'copy':
                text = f"{i}. {ctx_icon} Copy (Ctrl+C)"
                if 'preview' in action:
                    text += f" → '{action['preview']}'"
                text += "\n"
            elif action_type == 'paste':
                text = f"{i}. {ctx_icon} Paste (Ctrl+V)"
                if action.get('dynamic'):
                    text += " [Dynamic]"
                if 'preview' in action:
                    text += f" → '{action.get('preview', 'current clipboard')}'"
                text += "\n"
            elif action_type == 'cut':
                text = f"{i}. {ctx_icon} Cut (Ctrl+X)"
                if 'preview' in action:
                    text += f" → '{action['preview']}'"
                text += "\n"
            elif action_type == 'new_file':
                text = f"{i}. {ctx_icon} New File (Ctrl+N)\n"
            elif action_type == 'new_tab':
                text = f"{i}. {ctx_icon} New Tab (Ctrl+T)\n"
            elif action_type == 'select_all':
                text = f"{i}. {ctx_icon} Select All (Ctrl+A)\n"
            elif action_type == 'undo':
                text = f"{i}. {ctx_icon} Undo (Ctrl+Z)\n"
            elif action_type == 'redo':
                text = f"{i}. {ctx_icon} Redo (Ctrl+Y)\n"
            elif action_type == 'save':
                text = f"{i}. {ctx_icon} Save (Ctrl+S)\n"
            elif action_type == 'find':
                text = f"{i}. {ctx_icon} Find (Ctrl+F)\n"
            elif action_type == 'key':
                modifiers = []
                if action.get('ctrl'):
                    modifiers.append('Ctrl')
                if action.get('shift'):
                    modifiers.append('Shift')
                if action.get('alt'):
                    modifiers.append('Alt')
                    
                if modifiers:
                    text = f"{i}. {ctx_icon} Key: {'+'.join(modifiers)}+{action['key']}\n"
                else:
                    text = f"{i}. {ctx_icon} Key: {action['key']}\n"
            else:
                text = f"{i}. {ctx_icon} {action_type}\n"
                
            self.sequence_text.insert(tk.END, text)
            
        self.sequence_text.see(tk.END)

def main():
    root = tk.Tk()
    
    missing = []
    try:
        import pyperclip
    except ImportError:
        missing.append("pyperclip")
    try:
        import pynput
    except ImportError:
        missing.append("pynput")
    try:
        from PIL import ImageGrab
    except ImportError:
        missing.append("pillow")
    
    if missing:
        messagebox.showwarning("Missing Dependencies", 
                           f"Install: pip install {' '.join(missing)}")
    

    try:
        import pygetwindow
    except ImportError:
        messagebox.showinfo("Optional Dependency", 
                          "For enhanced window detection, install:\npip install pygetwindow")
    
    app = AdvancedRecorderGUI(root)
    root.mainloop() 

if __name__ == "__main__":
    main()
