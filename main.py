import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import pyautogui
import pyperclip
import time
import threading
import json
from pynput import mouse, keyboard
import queue
import pytesseract
from PIL import ImageGrab
import cv2
import numpy as np
import re
from collections import deque
from difflib import SequenceMatcher

class AdvancedRecorderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AutomatePro - Smart Context-Aware Mode")
        self.root.geometry("900x900")
        
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
        
        # Smart Context Detection
        self.current_context = "mixed"  # "excel", "absolute", "mixed"
        self.last_arrow_key_time = 0
        self.arrow_key_count = 0
        self.last_copy_paste_time = 0
        self.context_switch_times = deque(maxlen=5)
        self.recent_actions = deque(maxlen=20)
        
        self.setup_gui()
        
    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="AutomatePro - Dynamic Clipboard Mode", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=4, pady=10)
        
        # Context indicator
        context_frame = ttk.Frame(main_frame)
        context_frame.grid(row=1, column=0, columnspan=4, pady=5)
        ttk.Label(context_frame, text="Current Context:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        self.context_label = ttk.Label(context_frame, text="MIXED", 
                                       font=('Arial', 10), foreground='blue')
        self.context_label.pack(side=tk.LEFT, padx=5)
        
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
        
        options_frame = ttk.LabelFrame(main_frame, text="Smart Detection Options", padding="10")
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
        
        self.smart_context_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="🧠 Auto-Detect Context (Excel/Absolute)", 
                       variable=self.smart_context_var,
                       command=self.toggle_smart_mode).grid(row=1, column=0, padx=5, columnspan=2)
        
        self.record_clipboard_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Clipboard", 
                       variable=self.record_clipboard_var).grid(row=1, column=2, padx=5)
        
        # Context detection sensitivity
        sensitivity_frame = ttk.Frame(options_frame)
        sensitivity_frame.grid(row=2, column=0, columnspan=3, pady=5)
        ttk.Label(sensitivity_frame, text="Excel Detection Sensitivity:").pack(side=tk.LEFT, padx=5)
        self.sensitivity_var = tk.IntVar(value=2)
        ttk.Scale(sensitivity_frame, from_=1, to=5, variable=self.sensitivity_var, 
                 orient=tk.HORIZONTAL, length=150).pack(side=tk.LEFT, padx=5)
        ttk.Label(sensitivity_frame, text="(1=Low, 5=High)").pack(side=tk.LEFT)
        
        file_frame = ttk.LabelFrame(main_frame, text="File Operations", padding="10")
        file_frame.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
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
        
        sequence_frame = ttk.LabelFrame(main_frame, text="Sequence Logs", padding="10")
        sequence_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.sequence_text = scrolledtext.ScrolledText(sequence_frame, width=70, height=12)
        self.sequence_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Debug log
        debug_frame = ttk.LabelFrame(main_frame, text="Context Detection Log", padding="10")
        debug_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        self.debug_text = scrolledtext.ScrolledText(debug_frame, width=70, height=5)
        self.debug_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.status_var = tk.StringVar(value="Ready - Dynamic Clipboard Mode ENABLED")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        instructions = ttk.Label(main_frame, 
                                text="Developed by Nader Mahbub Khan\n" +
                                     "Press F9 to stop • Developed by Nader for Chandan Chakraborty",
                                font=('Arial', 9), foreground='gray')
        instructions.grid(row=8, column=0, columnspan=4, pady=5)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        sequence_frame.columnconfigure(0, weight=1)
        sequence_frame.rowconfigure(0, weight=1)
        debug_frame.columnconfigure(0, weight=1)
        
        self.setup_global_hotkey()
    
    def toggle_smart_mode(self):
        if self.smart_context_var.get():
            self.status_var.set("Smart Context Detection ENABLED")
            self.log_debug("🧠 Smart detection ON")
        else:
            self.status_var.set("Smart Context Detection DISABLED - Manual mode")
            self.log_debug("Manual mode activated")
    
    def detect_context(self, action_type, key=None):
        """Smart context detection based on recent actions"""
        if not self.smart_context_var.get():
            return "mixed"
        
        current_time = time.time()
        sensitivity = self.sensitivity_var.get()
        
        # Add to recent actions
        self.recent_actions.append({
            'type': action_type,
            'key': key,
            'time': current_time
        })
        
        # Count recent arrow keys (last 5 seconds)
        recent_arrows = sum(1 for a in self.recent_actions 
                           if a['time'] > current_time - 5 
                           and a.get('key') in ['up', 'down', 'left', 'right'])
        
        # Count recent copy/paste
        recent_copy_paste = sum(1 for a in self.recent_actions 
                               if a['time'] > current_time - 3 
                               and a['type'] in ['copy', 'paste'])
        
        # Count mouse clicks
        recent_clicks = sum(1 for a in self.recent_actions 
                           if a['time'] > current_time - 5 
                           and a['type'] in ['click', 'right_click'])
        
        # Excel detection logic
        excel_score = 0
        
        # Strong Excel indicators
        if recent_arrows >= sensitivity:
            excel_score += 3
        if recent_copy_paste > 0 and recent_arrows > 0:
            excel_score += 2
        if action_type in ['copy', 'paste'] and (current_time - self.last_arrow_key_time) < 2:
            excel_score += 2
        
        # Absolute mode indicators
        absolute_score = 0
        
        if key == 'n' and self.ctrl_pressed:  # Ctrl+N (new file)
            absolute_score += 5
            self.log_debug("🆕 Ctrl+N detected → ABSOLUTE mode")
        if key == 't' and self.ctrl_pressed:  # Ctrl+T (new tab)
            absolute_score += 5
            self.log_debug("🌐 Ctrl+T detected → BROWSER mode")
        if recent_clicks > sensitivity:
            absolute_score += 2
        if action_type == 'click':
            absolute_score += 1
        
        # Determine context
        if excel_score > absolute_score and excel_score >= 3:
            new_context = "excel"
            color = "green"
            symbol = "📊"
        elif absolute_score > excel_score and absolute_score >= 3:
            new_context = "absolute"
            color = "orange"
            symbol = "🎯"
        else:
            new_context = "mixed"
            color = "blue"
            symbol = "🔀"
        
        # Log context change
        if new_context != self.current_context:
            self.current_context = new_context
            self.context_switch_times.append(current_time)
            self.log_debug(f"{symbol} Context switched to: {new_context.upper()} (Excel:{excel_score}, Abs:{absolute_score})")
            self.root.after(0, lambda: self.context_label.config(
                text=new_context.upper(), 
                foreground=color
            ))
        
        return self.current_context
    
    def log_debug(self, message):
        """Log to debug window"""
        timestamp = time.strftime('%H:%M:%S')
        self.debug_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.debug_text.see(tk.END)
        # Keep only last 50 lines
        lines = self.debug_text.get('1.0', tk.END).split('\n')
        if len(lines) > 50:
            self.debug_text.delete('1.0', '2.0')
    
    def log_error(self, context, error):
        error_msg = f"[{time.strftime('%H:%M:%S')}] {context}: {str(error)}"
        self.error_log.append(error_msg)
        if self.debug_mode:
            print(error_msg)
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
        
    def get_clipboard_safe(self, max_retries=5, delay=0.1):
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
            time.sleep(0.2)
            
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
        
        self.status_var.set(f"🔴 Recording... (Dynamic Mode ON) - Press F9 to stop")
        
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
        self.pending_combo_action = None
        self.copy_action_id = 0
        self.current_context = "mixed"
        self.last_arrow_key_time = 0
        self.arrow_key_count = 0
        self.recent_actions.clear()
        
        self.last_clipboard = self.get_clipboard_safe()
        self.clipboard_before_copy = self.last_clipboard
        
        self.log_debug("🎬 Recording started - Dynamic clipboard mode active")
        
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
                
                # Detect context
                context = self.detect_context('click')
                
                # Always record right/middle clicks with absolute coords
                if button in [mouse.Button.right, mouse.Button.middle]:
                    action = {
                        'type': 'right_click' if button == mouse.Button.right else 'middle_click',
                        'x': x,
                        'y': y,
                        'time': current_time,
                        'context': 'absolute'  # Always absolute for special clicks
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
                        # Only record selection if in absolute mode
                        if context != "excel":
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
                        # Single click - only record if not in Excel mode
                        if context != "excel":
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
            
            # Only record scroll if not in Excel mode
            if context != "excel":
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

            # Track arrow keys for Excel detection
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

                    # Just capture a preview for display, NOT for playback
                    def delayed_capture(copy_id):
                        time.sleep(1.2)
                        new_clip = self.get_clipboard_safe()
                        if new_clip != self.clipboard_before_copy:
                            with self.action_lock:
                                for act in reversed(self.recorded_actions):
                                    if act.get('_id') == copy_id:
                                        # Only store preview for display
                                        preview = new_clip[:50] + "..." if len(new_clip) > 50 else new_clip
                                        act['preview'] = preview
                                        break
                            self.root.after(0, self.update_sequence_display)

                    threading.Thread(target=delayed_capture, args=(current_id,), daemon=True).start()
                    return

                elif vk == 86:  # 'V' - Paste
                    self.last_copy_paste_time = current_time
                    context = self.detect_context('paste')
                    
                    # ALWAYS use current clipboard for ALL contexts
                    action = {
                        'type': 'paste',
                        'time': current_time,
                        'context': context,
                        'dynamic': True  # Flag to indicate dynamic clipboard use
                    }
                    
                    self.add_action(action)
                    self.log_debug(f"📋 Paste recorded - will use CURRENT clipboard on playback")
                    return

                elif vk == 88:  # 'X' - Cut
                    self.clipboard_before_copy = self.get_clipboard_safe()
                    context = self.detect_context('cut')
                    action = {'type': 'cut', 'time': current_time, 'context': context}
                    self.add_action(action)

                    def capture_cut():
                        time.sleep(1.0)
                        new_clip = self.get_clipboard_safe()
                        if new_clip != self.clipboard_before_copy:
                            with self.action_lock:
                                if self.recorded_actions and self.recorded_actions[-1]['type'] == 'cut':
                                    preview = new_clip[:50] + "..." if len(new_clip) > 50 else new_clip
                                    self.recorded_actions[-1]['preview'] = preview
                            self.root.after(0, self.update_sequence_display)

                    threading.Thread(target=capture_cut, daemon=True).start()
                    return
                
                elif vk == 78:  # 'N' - Ctrl+N (New file/window)
                    context = self.detect_context('key', 'n')
                    action = {'type': 'new_file', 'time': current_time, 'context': context}
                    self.add_action(action)
                    self.log_debug("📄 Ctrl+N (New File) detected")
                    return
                
                elif vk == 84:  # 'T' - Ctrl+T (New tab)
                    context = self.detect_context('key', 't')
                    action = {'type': 'new_tab', 'time': current_time, 'context': context}
                    self.add_action(action)
                    self.log_debug("🌐 Ctrl+T (New Tab) detected")
                    return

                elif vk == 65:  # 'A'
                    action = {'type': 'select_all', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 90:  # 'Z'
                    action = {'type': 'undo', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 89:  # 'Y'
                    action = {'type': 'redo', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 83:  # 'S'
                    action = {'type': 'save', 'time': current_time}
                    self.add_action(action)
                    return
                elif vk == 70:  # 'F'
                    action = {'type': 'find', 'time': current_time}
                    self.add_action(action)
                    return

            # Regular key handling
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
                            context = action.get('context', 'mixed')
                            
                            def safe_coords(x, y):
                                x = max(0, min(screen_width - 1, x))
                                y = max(0, min(screen_height - 1, y))
                                return int(x), int(y)
                            
                            # Only execute mouse actions if not in Excel context
                            if action_type == 'click' and context != 'excel':
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
                                
                            elif action_type == 'selection' and context != 'excel':
                                start_x, start_y = safe_coords(action['start_x'], action['start_y'])
                                end_x, end_y = safe_coords(action['end_x'], action['end_y'])
                                pyautogui.moveTo(start_x, start_y, duration=0.2, tween=pyautogui.easeInOutQuad)
                                pyautogui.mouseDown()
                                time.sleep(0.05)
                                pyautogui.moveTo(end_x, end_y, duration=0.3, tween=pyautogui.easeInOutQuad)
                                time.sleep(0.05)
                                pyautogui.mouseUp()
                                self.safe_update_status("Playing: Selection/Drag")
                                
                            elif action_type == 'scroll' and context != 'excel':
                                x, y = safe_coords(action['x'], action['y'])
                                pyautogui.moveTo(x, y, duration=0.1)
                                scroll_amount = int(action['dy'] * 120)
                                pyautogui.scroll(scroll_amount)
                                self.safe_update_status("Playing: Scroll")
                                time.sleep(0.05)
                                
                            elif action_type == 'copy':
                                pyautogui.hotkey('ctrl', 'c')
                                ctx_tag = f" [{context.upper()}]"
                                self.safe_update_status(f"Playing: Copy{ctx_tag}")
                                time.sleep(0.6)
                                
                            elif action_type == 'paste':
                                # ALWAYS use current clipboard content
                                current_clip = self.get_clipboard_safe()
                                preview = current_clip[:30] + "..." if len(current_clip) > 30 else current_clip
                                self.safe_update_status(f"Playing: Paste [DYNAMIC: {preview}]")
                                pyautogui.hotkey('ctrl', 'v')
                                time.sleep(0.3)
                                
                            elif action_type == 'cut':
                                pyautogui.hotkey('ctrl', 'x')
                                self.safe_update_status("Playing: Cut (Ctrl+X)")
                                time.sleep(0.5)
                            
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
                                
                                # Extra delay for arrow keys in Excel context
                                if key in ['up', 'down', 'left', 'right'] and context == 'excel':
                                    time.sleep(0.25)
                                
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
        self.current_context = "mixed"
        self.context_label.config(text="MIXED", foreground="blue")
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
            context = action.get('context', 'mixed')
            
            # Context icon
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
                text = f"{i}. {ctx_icon} Paste (Ctrl+V) [DYNAMIC - uses current clipboard]\n"
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
        import pytesseract
    except ImportError:
        missing.append("pytesseract")
    try:
        import cv2
    except ImportError:
        missing.append("opencv-python")
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
    
    app = AdvancedRecorderGUI(root)
    root.mainloop() 

if __name__ == "__main__":
    main()
