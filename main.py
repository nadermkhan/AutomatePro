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

# Optional: Set Tesseract path if not in PATH (Windows example)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class AdvancedRecorderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AutomatePro for Chandan Chakraborty")
        self.root.geometry("900x800")
        
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
        self.last_copy_index = -1
        
        self.max_actions = 10000
        
        self.debug_mode = False
        self.error_log = []
        
        self.setup_gui()
        
    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="AutomatePro", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        control_frame = ttk.LabelFrame(main_frame, text="Controls", padding="10")
        control_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.record_btn = ttk.Button(control_frame, text="Start Recording", 
                                     command=self.toggle_recording, width=20)
        self.record_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.play_btn = ttk.Button(control_frame, text="Play", 
                                   command=self.play_sequence, width=20)
        self.play_btn.grid(row=0, column=1, padx=5, pady=5)
        
        self.clear_btn = ttk.Button(control_frame, text="Clear", 
                                    command=self.clear_sequence, width=20)
        self.clear_btn.grid(row=0, column=2, padx=5, pady=5)
        
        options_frame = ttk.LabelFrame(main_frame, text="Recording Options", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.record_keyboard_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Keyboard", 
                       variable=self.record_keyboard_var).grid(row=0, column=0, padx=5)
        
        self.record_mouse_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Mouse", 
                       variable=self.record_mouse_var).grid(row=0, column=1, padx=5)
        
        self.record_scroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="MouseScroll", 
                       variable=self.record_scroll_var).grid(row=0, column=2, padx=5)
        
        self.use_ocr_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="OCR (Context Awareness)", 
                       variable=self.use_ocr_var).grid(row=1, column=0, padx=5)
        
        self.record_clipboard_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Clipboard", 
                       variable=self.record_clipboard_var).grid(row=1, column=1, padx=5)
        
        self.smart_selection_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Smart Selection", 
                       variable=self.smart_selection_var).grid(row=1, column=2, padx=5)
        
        file_frame = ttk.LabelFrame(main_frame, text="File Operations", padding="10")
        file_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
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
        sequence_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.sequence_text = scrolledtext.ScrolledText(sequence_frame, width=70, height=15)
        self.sequence_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        instructions = ttk.Label(main_frame, 
                                text="Press F9 to stop recording/playing • F10 to pause/resume playback\n" +
                                     "Developed by Nader for Chandan Chakraborty",
                                font=('Arial', 9), foreground='gray')
        instructions.grid(row=6, column=0, columnspan=3, pady=5)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        sequence_frame.columnconfigure(0, weight=1)
        sequence_frame.rowconfigure(0, weight=1)
        
        self.setup_global_hotkey()
    
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
    
    def analyze_selection_context(self, clipboard_content):
        if not clipboard_content:
            return None
            
        words = clipboard_content.split()
        if len(words) < 3:
            return {
                'full_text': clipboard_content,
                'start_context': clipboard_content,
                'end_context': clipboard_content,
                'pattern': clipboard_content,
                'words_count': len(words)
            }
        
        start_words = ' '.join(words[:min(5, len(words)//2)])
        end_words = ' '.join(words[max(-5, -len(words)//2):])
        
        return {
            'full_text': clipboard_content,
            'start_context': start_words,
            'end_context': end_words,
            'start_pattern': re.escape(start_words),
            'end_pattern': re.escape(end_words),
            'words_count': len(words)
        }
            
    def capture_text_around_cursor(self, x, y, radius=50):
        if not self.use_ocr_var.get():
            return None
            
        try:
            x1, y1 = max(0, x - radius), max(0, y - radius)
            x2, y2 = x + radius, y + radius
            
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            img_np = np.array(screenshot)
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            text = pytesseract.image_to_string(thresh, config='--psm 8 -c tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ .,')
            
            return text.strip() if text.strip() else None
        except Exception as e:
            self.log_error("OCR around cursor", e)
            return None

    def capture_text_at_selection(self, x1, y1, x2, y2):
        if not self.use_ocr_var.get():
            return None
            
        try:
            x1, x2 = min(x1, x2), max(x1, x2)
            y1, y2 = min(y1, y2), max(y1, y2)
            
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            img_np = np.array(screenshot)
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            text = pytesseract.image_to_string(thresh, config='--psm 6')
            
            return text.strip() if text.strip() else None
        except Exception as e:
            self.log_error("OCR capture", e)
            return None

    def _fuzzy_match(self, a, b, threshold=0.8):
        if not a or not b:
            return False
        if a == b:
            return True
        ratio = SequenceMatcher(None, a.lower(), b.lower()).ratio()
        return ratio >= threshold

    def locate_text_on_screen(self, search_text, min_confidence=50):
        if not search_text or len(search_text.strip()) < 2:
            return None
            
        try:
            screenshot = ImageGrab.grab()
            img_np = np.array(screenshot)
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

            data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
            n_boxes = len(data['text'])
            search_words = [w for w in search_text.split() if w]

            if not search_words:
                return None

            # Try multi-word phrase first
            for i in range(n_boxes - len(search_words) + 1):
                matched = True
                phrase_boxes = []
                for j, word in enumerate(search_words):
                    idx = i + j
                    conf = int(data['conf'][idx])
                    detected = data['text'][idx].strip()
                    if conf < min_confidence or not detected:
                        matched = False
                        break
                    if not self._fuzzy_match(word, detected, threshold=0.75):
                        matched = False
                        break
                    phrase_boxes.append(idx)

                if matched and phrase_boxes:
                    # Calculate bounding box of full phrase
                    lefts = [data['left'][idx] for idx in phrase_boxes]
                    tops = [data['top'][idx] for idx in phrase_boxes]
                    widths = [data['width'][idx] for idx in phrase_boxes]
                    heights = [data['height'][idx] for idx in phrase_boxes]
                    
                    x1 = min(lefts)
                    y1 = min(tops)
                    x2 = max(lefts[i] + widths[i] for i in range(len(lefts)))
                    y2 = max(tops[i] + heights[i] for i in range(len(tops)))
                    
                    x_center = (x1 + x2) // 2
                    y_center = (y1 + y2) // 2
                    return (x_center, y_center)

            # Fallback: single word (first word)
            first_word = search_words[0]
            for i in range(n_boxes):
                conf = int(data['conf'][i])
                detected = data['text'][i].strip()
                if conf >= min_confidence and detected and self._fuzzy_match(first_word, detected, threshold=0.8):
                    x = data['left'][i] + data['width'][i] // 2
                    y = data['top'][i] + data['height'][i] // 2
                    return (x, y)

            return None

        except Exception as e:
            self.log_error("Text location failed", e)
            return None

    def navigate_to_context_silently(self, context_info):
        if not context_info or not self.use_ocr_var.get():
            return False
            
        start_context = context_info.get('start_context', '').strip()
        if not start_context:
            return False
            
        location = self.locate_text_on_screen(start_context)
        if location:
            x, y = location
            screen_w, screen_h = pyautogui.size()
            x = max(10, min(screen_w - 10, x))
            y = max(10, min(screen_h - 10, y))
            pyautogui.moveTo(x, y, duration=0.3, tween=pyautogui.easeInOutQuad)
            time.sleep(0.2)
            return True
        return False
            
    def toggle_recording(self):
        if not self.recording:
            self.start_recording()
        else:
            self.stop_recording()
    
    def start_recording(self):
        self.recording = True
        self.record_btn.config(text="⏹ Stop Recording")
        self.play_btn.config(state='disabled')
        self.clear_btn.config(state='disabled')
        self.save_btn.config(state='disabled')
        self.load_btn.config(state='disabled')
        self.status_var.set("Recording... (Press F9 to stop)")
        
        with self.action_lock:
            self.recorded_actions = []
        self.update_sequence_display()
        
        self.mouse_pressed = False
        self.selection_start = None
        self.selection_end = None
        self.is_dragging = False
        self.drag_start_time = None
        self.scroll_before_selection = []
        self.context_actions = []
        self.current_combo = set()
        self.ctrl_pressed = False
        self.shift_pressed = False
        self.alt_pressed = False
        self.combo_detected = False
        self.last_copy_index = -1
        self.selection_context = {}
        self.clipboard_history.clear()
        self.pending_combo_action = None
        
        self.last_clipboard = self.get_clipboard_safe()
        self.clipboard_before_copy = self.last_clipboard
        
        clipboard_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
        clipboard_thread.start()
        
        def on_click(x, y, button, pressed):
            if not self.recording or not self.record_mouse_var.get():
                return
                
            if pressed:
                self.mouse_pressed = True
                self.selection_start = (x, y, time.time())
                self.drag_start_time = time.time()
                
                if self.smart_selection_var.get():
                    current_time = time.time()
                    with self.action_lock:
                        self.context_actions = [
                            action for action in self.recorded_actions[-20:]
                            if current_time - action.get('time', 0) < 5
                        ]
                
                if button == mouse.Button.right:
                    action = {
                        'type': 'right_click',
                        'x': x,
                        'y': y,
                        'time': time.time()
                    }
                    if self.use_ocr_var.get():
                        text = self.capture_text_around_cursor(x, y)
                        if text:
                            action['element_text'] = text
                    self.add_action(action)
                elif button == mouse.Button.middle:
                    action = {
                        'type': 'middle_click',
                        'x': x,
                        'y': y,
                        'time': time.time()
                    }
                    if self.use_ocr_var.get():
                        text = self.capture_text_around_cursor(x, y)
                        if text:
                            action['element_text'] = text
                    self.add_action(action)
                    
            else:
                if self.mouse_pressed and self.selection_start:
                    self.selection_end = (x, y, time.time())
                    
                    start_x, start_y, start_time = self.selection_start
                    distance = ((x - start_x)**2 + (y - start_y)**2)**0.5
                    duration = time.time() - self.drag_start_time
                    
                    if distance > 10 and duration > 0.1:
                        recent_copy = None
                        with self.action_lock:
                            for i in range(len(self.recorded_actions) - 1, max(0, len(self.recorded_actions) - 10), -1):
                                if self.recorded_actions[i].get('type') == 'copy':
                                    recent_copy = i
                                    break
                        
                        action = {
                            'type': 'text_selection',
                            'start_x': start_x,
                            'start_y': start_y,
                            'end_x': x,
                            'end_y': y,
                            'duration': duration,
                            'time': time.time()
                        }
                        
                        if recent_copy is not None:
                            action['linked_copy_index'] = recent_copy
                        
                        if self.smart_selection_var.get() and self.context_actions:
                            scroll_actions = [a for a in self.context_actions if a.get('type') == 'scroll']
                            if scroll_actions:
                                action['context_scrolls'] = scroll_actions
                        
                        if self.use_ocr_var.get():
                            selected_text = self.capture_text_at_selection(start_x, start_y, x, y)
                            if selected_text:
                                action['selected_text'] = selected_text
                        
                        self.add_action(action)
                    else:
                        action = {
                            'type': 'click',
                            'x': start_x,
                            'y': start_y,
                            'button': str(button),
                            'time': time.time()
                        }
                        if self.use_ocr_var.get():
                            text = self.capture_text_around_cursor(start_x, start_y)
                            if text:
                                action['element_text'] = text
                        self.add_action(action)
                
                self.mouse_pressed = False
                self.selection_start = None
                self.is_dragging = False
                
        def on_move(x, y):
            if not self.recording or not self.record_mouse_var.get():
                return
            
            if self.mouse_pressed and self.selection_start:
                start_x, start_y, _ = self.selection_start
                distance = ((x - start_x)**2 + (y - start_y)**2)**0.5
                if distance > 5:
                    self.is_dragging = True
                
        def on_scroll(x, y, dx, dy):
            if not self.recording or not self.record_scroll_var.get():
                return
                
            current_time = time.time()
            if current_time - self.last_scroll_time < 0.05:
                return
            self.last_scroll_time = current_time
            
            action = {
                'type': 'scroll',
                'x': x,
                'y': y,
                'dx': dx,
                'dy': dy,
                'time': current_time
            }
            self.add_action(action)
            
        def on_key_press(key):
            if not self.recording or not self.record_keyboard_var.get():
                return
            if key == keyboard.Key.f9:
                return

            vk = getattr(key, 'vk', None)

            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, keyboard.Key.ctrl]:
                self.ctrl_pressed = True
                return
            elif key in [keyboard.Key.shift_l, keyboard.Key.shift_r, keyboard.Key.shift]:
                self.shift_pressed = True
                return
            elif key in [keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt, keyboard.Key.alt_gr]:
                self.alt_pressed = True
                return

            if self.ctrl_pressed and vk is not None:
                current_time = time.time()
                if vk == 67:  # 'C'
                    self.clipboard_before_copy = self.get_clipboard_safe()
                    action = {'type': 'copy', 'time': current_time}
                    self.add_action(action)
                    self.last_copy_index = len(self.recorded_actions) - 1

                    def delayed_capture():
                        time.sleep(0.6)
                        new_clip = self.get_clipboard_safe()
                        if new_clip != self.clipboard_before_copy:
                            context = self.analyze_selection_context(new_clip)
                            with self.action_lock:
                                if 0 <= self.last_copy_index < len(self.recorded_actions):
                                    self.recorded_actions[self.last_copy_index].update({
                                        'copied_content': new_clip,
                                        'selection_context': context
                                    })
                                    for i in range(self.last_copy_index - 1, max(-1, self.last_copy_index - 6), -1):
                                        if self.recorded_actions[i].get('type') == 'text_selection':
                                            self.recorded_actions[i].update({
                                                'copied_content': new_clip,
                                                'selection_context': context
                                            })
                                            break
                            self.root.after(0, self.update_sequence_display)

                    threading.Thread(target=delayed_capture, daemon=True).start()
                    return

                elif vk == 86:  # 'V'
                    clipboard_content = self.get_clipboard_safe()
                    action = {
                        'type': 'paste',
                        'time': current_time,
                        'clipboard_content': clipboard_content
                    }
                    self.add_action(action)
                    return

                elif vk == 88:  # 'X'
                    self.clipboard_before_copy = self.get_clipboard_safe()
                    action = {'type': 'cut', 'time': current_time}
                    self.add_action(action)

                    def capture_cut():
                        time.sleep(0.6)
                        new_clip = self.get_clipboard_safe()
                        if new_clip != self.clipboard_before_copy:
                            with self.action_lock:
                                if self.recorded_actions and self.recorded_actions[-1]['type'] == 'cut':
                                    self.recorded_actions[-1]['cut_content'] = new_clip
                            self.root.after(0, self.update_sequence_display)

                    threading.Thread(target=capture_cut, daemon=True).start()
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

            key_repr = None
            if hasattr(key, 'char') and key.char and key.char.isprintable():
                key_repr = key.char
            elif hasattr(key, 'name'):
                key_repr = key.name
            elif vk is not None:
                if 65 <= vk <= 90:
                    key_repr = chr(vk).lower()
                elif 48 <= vk <= 57:
                    key_repr = chr(vk)

            if key_repr:
                action = {
                    'type': 'key',
                    'key': key_repr,
                    'ctrl': self.ctrl_pressed,
                    'shift': self.shift_pressed,
                    'alt': self.alt_pressed,
                    'time': time.time()
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
            on_scroll=on_scroll,
            on_move=on_move
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
        self.clear_btn.config(state='normal')
        self.save_btn.config(state='normal')
        self.load_btn.config(state='normal')
        
        if self.mouse_listener:
            try:
                self.mouse_listener.stop()
            except:
                pass
        if self.keyboard_listener:
            try:
                self.keyboard_listener.stop()
            except:
                pass
        
        with self.action_lock:
            action_count = len(self.recorded_actions)
        self.status_var.set(f"Recording stopped. {action_count} actions recorded.")
    
    def play_sequence(self):
        with self.action_lock:
            if not self.recorded_actions:
                messagebox.showwarning("Warning", "No actions recorded")
                return
            actions_to_play = self.recorded_actions.copy()
            
        def play_thread():
            self.playing = True
            self.play_btn.config(state='disabled')
            self.record_btn.config(state='disabled')
            self.clear_btn.config(state='disabled')
            self.save_btn.config(state='disabled')
            self.load_btn.config(state='disabled')
            
            self.safe_update_status("Playing sequence... (Press F9 to stop)")
            
            time.sleep(2)
            
            speed = self.speed_var.get()
            prev_time = None
            
            screen_width, screen_height = pyautogui.size()
            
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
                    
                    def safe_coords(x, y):
                        x = max(0, min(screen_width - 1, x))
                        y = max(0, min(screen_height - 1, y))
                        return int(x), int(y)
                    
                    if action_type == 'click':
                        x, y = safe_coords(action['x'], action['y'])
                        pyautogui.moveTo(x, y, duration=0.2, tween=pyautogui.easeInOutQuad)
                        pyautogui.click()
                        text_info = f" '{action.get('element_text', '')}'" if 'element_text' in action else ""
                        self.safe_update_status(f"Playing: Click{text_info} at ({x}, {y})")
                        
                    elif action_type == 'right_click':
                        x, y = safe_coords(action['x'], action['y'])
                        pyautogui.moveTo(x, y, duration=0.2, tween=pyautogui.easeInOutQuad)
                        pyautogui.rightClick()
                        text_info = f" '{action.get('element_text', '')}'" if 'element_text' in action else ""
                        self.safe_update_status(f"Playing: Right-click{text_info} at ({x}, {y})")
                        
                    elif action_type == 'middle_click':
                        x, y = safe_coords(action['x'], action['y'])
                        pyautogui.moveTo(x, y, duration=0.2, tween=pyautogui.easeInOutQuad)
                        pyautogui.middleClick()
                        text_info = f" '{action.get('element_text', '')}'" if 'element_text' in action else ""
                        self.safe_update_status(f"Playing: Middle-click{text_info} at ({x}, {y})")
                        
                    elif action_type == 'text_selection':
                        start_x, start_y = safe_coords(action['start_x'], action['start_y'])
                        end_x, end_y = safe_coords(action['end_x'], action['end_y'])
                        
                        if self.smart_selection_var.get() and 'selection_context' in action:
                            context = action['selection_context']
                            self.safe_update_status("Locating context silently...")
                            if context and 'start_context' in context:
                                found = self.navigate_to_context_silently(context)
                                if found:
                                    self.safe_update_status("Context located")
                            time.sleep(0.3)
                            
                            if context and context.get('words_count', 0) > 20:
                                pyautogui.click(start_x, start_y)
                                time.sleep(0.1)
                                pyautogui.click(clicks=3, interval=0.1)
                                time.sleep(0.3)
                            else:
                                pyautogui.moveTo(start_x, start_y, duration=0.2, tween=pyautogui.easeInOutQuad)
                                pyautogui.mouseDown()
                                time.sleep(0.05)
                                pyautogui.moveTo(end_x, end_y, duration=0.3, tween=pyautogui.easeInOutQuad)
                                time.sleep(0.05)
                                pyautogui.mouseUp()
                                time.sleep(0.1)
                        else:
                            pyautogui.moveTo(start_x, start_y, duration=0.2, tween=pyautogui.easeInOutQuad)
                            pyautogui.mouseDown()
                            time.sleep(0.05)
                            pyautogui.moveTo(end_x, end_y, duration=0.3, tween=pyautogui.easeInOutQuad)
                            time.sleep(0.05)
                            pyautogui.mouseUp()
                        
                        self.safe_update_status("Playing: Text selection")
                        time.sleep(0.2)
                        
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
                        time.sleep(0.4)
                        
                    elif action_type == 'paste':
                        if 'clipboard_content' in action and action['clipboard_content']:
                            try:
                                pyperclip.copy(action['clipboard_content'])
                                time.sleep(0.2)
                            except Exception as e:
                                self.log_error("Set clipboard for paste", e)
                        pyautogui.hotkey('ctrl', 'v')
                        self.safe_update_status("Playing: Paste (Ctrl+V)")
                        time.sleep(0.3)
                        
                    elif action_type == 'cut':
                        pyautogui.hotkey('ctrl', 'x')
                        self.safe_update_status("Playing: Cut (Ctrl+X)")
                        time.sleep(0.3)
                        
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
                        
                    elif action_type == 'clipboard_change':
                        pyperclip.copy(action['content'])
                        self.safe_update_status("Playing: Clipboard updated")
                        time.sleep(0.1)
                        
                    elif action_type == 'key':
                        key = action['key']
                        if key.startswith('Key.'):
                            key_name = key.replace('Key.', '')
                            if key_name in ['space', 'enter', 'tab', 'backspace', 'delete', 
                                          'home', 'end', 'page_up', 'page_down', 'up', 'down', 
                                          'left', 'right', 'escape', 'esc']:
                                if key_name == 'esc':
                                    key_name = 'escape'
                                pyautogui.press(key_name)
                        else:
                            modifiers = []
                            if action.get('ctrl'):
                                modifiers.append('ctrl')
                            if action.get('shift'):
                                modifiers.append('shift')
                            if action.get('alt'):
                                modifiers.append('alt')
                            
                            if modifiers:
                                pyautogui.hotkey(*modifiers, key)
                            else:
                                pyautogui.write(key)
                        
                        self.safe_update_status(f"Playing: Key press '{key}'")
                        
                except Exception as e:
                    self.log_error(f"Playing action {i}", e)
                    
            self.playing = False
            self.play_btn.config(state='normal')
            self.record_btn.config(state='normal')
            self.clear_btn.config(state='normal')
            self.save_btn.config(state='normal')
            self.load_btn.config(state='normal')
            self.safe_update_status("Playback completed")
            
        threading.Thread(target=play_thread, daemon=True).start()
        
    def clear_sequence(self):
        with self.action_lock:
            self.recorded_actions = []
        self.sequence_text.delete('1.0', tk.END)
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
            
            if action_type == 'click':
                text = f"{i}. Click at ({action['x']}, {action['y']})"
                if 'element_text' in action:
                    text += f" → '{action['element_text']}'"
                text += "\n"
            elif action_type == 'right_click':
                text = f"{i}. Right-click at ({action['x']}, {action['y']})"
                if 'element_text' in action:
                    text += f" → '{action['element_text']}'"
                text += "\n"
            elif action_type == 'middle_click':
                text = f"{i}. Middle-click at ({action['x']}, {action['y']})"
                if 'element_text' in action:
                    text += f" → '{action['element_text']}'"
                text += "\n"
            elif action_type == 'text_selection':
                text = f"{i}. Text Selection from ({action['start_x']}, {action['start_y']}) to ({action['end_x']}, {action['end_y']})"
                if 'copied_content' in action:
                    preview = action['copied_content'][:30] + '...' if len(action['copied_content']) > 30 else action['copied_content']
                    text += f"\n    → Content: '{preview}'"
                if 'selection_context' in action:
                    context = action['selection_context']
                    text += f"\n    → Context: Start='{context.get('start_context', '')[:20]}...'"
                if 'selected_text' in action:
                    preview = action['selected_text'][:30] + '...' if len(action['selected_text']) > 30 else action['selected_text']
                    text += f"\n    → OCR: '{preview}'"
                text += "\n"
            elif action_type == 'selection':
                text = f"{i}. Selection/Drag from ({action['start_x']}, {action['start_y']}) to ({action['end_x']}, {action['end_y']})\n"
            elif action_type == 'scroll':
                direction = "down" if action['dy'] < 0 else "up"
                text = f"{i}. Scroll {direction} at ({action['x']}, {action['y']})\n"
            elif action_type == 'copy':
                text = f"{i}. Copy (Ctrl+C)"
                if 'copied_content' in action:
                    preview = action['copied_content'][:40] + '...' if len(action['copied_content']) > 40 else action['copied_content']
                    text += f" - Copied: '{preview}'"
                text += "\n"
            elif action_type == 'paste':
                text = f"{i}. Paste (Ctrl+V)"
                if 'clipboard_content' in action:
                    preview = action['clipboard_content'][:40] + '...' if len(action['clipboard_content']) > 40 else action['clipboard_content']
                    text += f" - Content: '{preview}'"
                text += "\n"
            elif action_type == 'cut':
                text = f"{i}. Cut (Ctrl+X)"
                if 'cut_content' in action:
                    preview = action['cut_content'][:40] + '...' if len(action['cut_content']) > 40 else action['cut_content']
                    text += f" - Cut: '{preview}'"
                text += "\n"
            elif action_type == 'select_all':
                text = f"{i}. Select All (Ctrl+A)\n"
            elif action_type == 'undo':
                text = f"{i}. Undo (Ctrl+Z)\n"
            elif action_type == 'redo':
                text = f"{i}. Redo (Ctrl+Y)\n"
            elif action_type == 'save':
                text = f"{i}. Save (Ctrl+S)\n"
            elif action_type == 'find':
                text = f"{i}. Find (Ctrl+F)\n"
            elif action_type == 'clipboard_change':
                content_preview = action['content'][:50] + '...' if len(action['content']) > 50 else action['content']
                text = f"{i}. Clipboard changed: '{content_preview}'\n"
            elif action_type == 'key':
                modifiers = []
                if action.get('ctrl'):
                    modifiers.append('Ctrl')
                if action.get('shift'):
                    modifiers.append('Shift')
                if action.get('alt'):
                    modifiers.append('Alt')
                    
                if modifiers:
                    text = f"{i}. Key press: {'+'.join(modifiers)}+{action['key']}\n"
                else:
                    text = f"{i}. Key press: {action['key']}\n"
            else:
                text = f"{i}. {action_type}\n"
                
            self.sequence_text.insert(tk.END, text)
            
        self.sequence_text.see(tk.END)

def main():
    root = tk.Tk()
    
    try:
        import pytesseract
        import cv2
        import pyperclip
    except ImportError as e:
        messagebox.showerror("Missing Dependencies", 
                           f"Please install required packages:\n"
                           f"pip install pytesseract opencv-python pyperclip pynput pillow\n\n"
                           f"For OCR support, also install Tesseract-OCR from:\n"
                           f"https://github.com/UB-Mannheim/tesseract/wiki")
        return
    
    app = AdvancedRecorderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
