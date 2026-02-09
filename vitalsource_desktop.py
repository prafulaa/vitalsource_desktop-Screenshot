"""
VitalSource Desktop App to PDF Converter (v2)
Captures pages directly from the VitalSource Bookshelf desktop application.
- User sets "Next Page" button location manually
- Press 'q' to stop at any time
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pathlib import Path
from PIL import Image, ImageGrab
import pyautogui
from fpdf import FPDF
import keyboard

# --- Configuration ---
SCRIPT_DIR = Path(__file__).resolve().parent
TEMP_IMAGE_DIR = SCRIPT_DIR / "temp_book_pages"

# Disable pyautogui failsafe (move mouse to corner to abort)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


class GlobalState:
    relative_offset = None  # (offset_x, offset_y) from window top-left
    paused = False


def find_vitalsource_window():
    """Find the VitalSource Bookshelf window."""
    try:
        import pygetwindow as gw
        windows = gw.getWindowsWithTitle('Bookshelf')
        if windows:
            return windows[0]
        # Try other possible titles
        for title in ['VitalSource', 'Vitalsource Bookshelf']:
            windows = gw.getWindowsWithTitle(title)
            if windows:
                return windows[0]
    except Exception:
        pass
    return None


def capture_window(window, path, crop_margins=True):
    """Capture a screenshot of the specified window."""
    try:
        # Get window position and size
        left, top, width, height = window.left, window.top, window.width, window.height
        
        # Capture the window area
        screenshot = ImageGrab.grab(bbox=(left, top, left + width, top + height))
        
        if crop_margins:
            # Crop to get just the book content (center area)
            w, h = screenshot.size
            # Remove: left sidebar (~280px), top bar (~80px), bottom bar (~50px), right margin (~50px)
            left_crop = 280
            top_crop = 80
            right_crop = w - 50
            bottom_crop = h - 50
            
            # Make sure crop is valid
            if right_crop > left_crop and bottom_crop > top_crop:
                screenshot = screenshot.crop((left_crop, top_crop, right_crop, bottom_crop))
        
        screenshot.save(str(path), optimize=True)
        return True
    except Exception as e:
        print(f"Capture error: {e}")
        return False


def click_next_page(window):
    """Click the user-defined next page coordinates."""
    if GlobalState.paused:
        print("Paused...")
        return False
        
    if GlobalState.relative_offset and window:
        try:
            abs_x = window.left + GlobalState.relative_offset[0]
            abs_y = window.top + GlobalState.relative_offset[1]
            pyautogui.click(abs_x, abs_y)
            return True
        except Exception:
            return False
    return False


def is_valid_image(path):
    """Check if an image file is valid."""
    try:
        with Image.open(path) as img:
            img.verify()
        return os.path.getsize(path) > 1000
    except Exception:
        return False


def create_pdf_from_images(image_files, output_pdf_path, log):
    """Assemble images into a PDF."""
    if not image_files:
        log("No images found to create a PDF.")
        return False

    log(f"Assembling {len(image_files)} pages into PDF...")
    pdf = FPDF(unit="pt")
    pdf.set_auto_page_break(auto=False)
    pdf.set_margins(0, 0, 0)

    for image_file in sorted(image_files):
        with Image.open(image_file) as img:
            width, height = img.size
        pdf.add_page(format=(width, height))
        pdf.image(str(image_file), 0, 0, width, height)

    pdf.output(output_pdf_path)
    log(f"PDF created: {output_pdf_path}")
    return True


def run_capture(total_pages, delay_ms, log, progress_cb, done_cb, stop_event):
    """Run the capture process."""
    if not GlobalState.relative_offset:
        log("ERROR: Next Button location not set!")
        log("Please click 'Set Next Button' first.")
        done_cb()
        return

    # Setup Hotkeys
    def _stop():
        log("Kill Switch (F10) detected!")
        stop_event.set()
        
    def _toggle_pause():
        GlobalState.paused = not GlobalState.paused
        state = "PAUSED" if GlobalState.paused else "RESUMED"
        log(f"Capture {state} via F9")

    try:
        keyboard.add_hotkey('f10', _stop)
        keyboard.add_hotkey('f9', _toggle_pause)
    except Exception as e:
        log(f"Warning: Could not register global hotkeys: {e}")

    captured_files = []
    
    # Create temp directory
    if not TEMP_IMAGE_DIR.exists():
        TEMP_IMAGE_DIR.mkdir()
    
    # Check for existing captures
    existing = sorted(TEMP_IMAGE_DIR.glob("page_*.png"))
    if existing:
        valid = [f for f in existing if is_valid_image(f)]
        if valid:
            log(f"Found {len(valid)} existing pages. Resuming...")
            captured_files = list(valid)
    
    # Find VitalSource window
    log("Looking for VitalSource Bookshelf window...")
    window = find_vitalsource_window()
    
    if window is None:
        log("=" * 50)
        log("ERROR: VitalSource Bookshelf window not found!")
        log("Please open VitalSource Bookshelf and your book.")
        log("=" * 50)
        done_cb()
        return
    
    log(f"Found window: {window.title}")
    
    # Bring window to front
    try:
        window.activate()
        time.sleep(0.5)
    except Exception:
        log("Could not activate window. Make sure it's visible.")
    
    log("=" * 50)
    log("Starting capture in 3 seconds...")
    log("Press 'q' or 'F10' to STOP at any time.")
    log("Press 'F9' to PAUSE/RESUME.")
    log("=" * 50)
    time.sleep(3)
    
    if stop_event.is_set():
        done_cb()
        return
    
    delay_sec = delay_ms / 1000.0
    start_time = time.time()
    start_page = len(captured_files) + 1
    
    log(f"Capturing from page {start_page}...")
    
    page_num = start_page - 1
    while not stop_event.is_set():
        if keyboard.is_pressed('q'): # Keeping 'q' as backup
            log("Stop key (q) detected!")
            stop_event.set()
            break
            
        if GlobalState.paused:
            time.sleep(0.5)
            continue

        page_num += 1
        
        if total_pages and page_num > total_pages:
            break
        
        screenshot_path = TEMP_IMAGE_DIR / f"page_{page_num:04d}.png"
        
        # Skip existing
        if screenshot_path.exists() and screenshot_path in captured_files:
            click_next_page(window)
            time.sleep(delay_sec)
            continue
        
        # Re-find window in case it moved
        window = find_vitalsource_window()
        if window is None:
            log("Window lost! Waiting...")
            time.sleep(2)
            window = find_vitalsource_window()
            if window is None:
                log("Window still not found. Stopping.")
                break
        
        # Capture
        if capture_window(window, screenshot_path):
            captured_files.append(screenshot_path)
            
            # Progress
            elapsed = time.time() - start_time
            pages_done = page_num - start_page + 1
            rate = pages_done / elapsed if elapsed > 0 else 0
            
            if total_pages:
                pct = int(page_num / total_pages * 100)
                remaining = (total_pages - page_num) / rate if rate > 0 else 0
                mins, secs = divmod(int(remaining), 60)
                log(f"Page {page_num}/{total_pages}  ({rate:.1f} p/s, ~{mins}m{secs:02d}s left)")
                progress_cb(pct)
            else:
                log(f"Page {page_num} captured  ({rate:.1f} p/s)")
                progress_cb(-1)
        else:
            log(f"Failed to capture page {page_num}")
        
        # Next page
        click_next_page(window)
        time.sleep(delay_sec)
    
    # Create PDF
    if stop_event.is_set():
        log("Capture stopped.")
    
    all_files = sorted(TEMP_IMAGE_DIR.glob("page_*.png"))
    valid_files = [f for f in all_files if is_valid_image(f)]
    
    if valid_files:
        pdf_path = SCRIPT_DIR / "converted_book.pdf"
        create_pdf_from_images(valid_files, pdf_path, log)
        
        # Cleanup
        if not stop_event.is_set(): # Only cleanup if finished fully
            log("Cleaning up temp images...")
            import shutil
            shutil.rmtree(TEMP_IMAGE_DIR)
    
    log("Done!")
    keyboard.remove_hotkey('f10')
    keyboard.remove_hotkey('f9')
    progress_cb(100)
    done_cb()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VitalSource Desktop Capture v2")
        self.geometry("600x580")
        self.resizable(False, False)
        self.configure(bg="#1e1e2e")
        
        self.stop_event = threading.Event()
        self.worker_thread = None
        
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        
        bg = "#1e1e2e"
        fg = "#cdd6f4"
        accent = "#89b4fa"
        entry_bg = "#313244"
        
        style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=bg, foreground=accent, font=("Segoe UI", 16, "bold"))
        style.configure("TEntry", fieldbackground=entry_bg, foreground=fg, font=("Segoe UI", 10))
        style.configure("TButton", background="#45475a", foreground=fg, font=("Segoe UI", 10, "bold"))
        style.configure("Accent.TButton", background=accent, foreground="#1e1e2e", font=("Segoe UI", 11, "bold"))
        style.configure("Stop.TButton", background="#f38ba8", foreground="#1e1e2e", font=("Segoe UI", 11, "bold"))
        style.configure("green.Horizontal.TProgressbar", troughcolor=entry_bg, background="#a6e3a1")
        
        # Header
        ttk.Label(self, text="VitalSource Desktop Capture v2", style="Header.TLabel").pack(pady=(18, 2))
        
        # Instructions
        instructions = tk.Frame(self, bg=bg)
        instructions.pack(padx=24, fill="x")
        
        tk.Label(instructions, text="Instructions:", bg=bg, fg=accent, font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")
        tk.Label(instructions, text="1. Click 'Set Next Button', hover over the '>' button, press 'n'.", bg=bg, fg=fg, font=("Segoe UI", 9), anchor="w").pack(fill="x")
        tk.Label(instructions, text="2. Enter total pages and click Start.", bg=bg, fg=fg, font=("Segoe UI", 9), anchor="w").pack(fill="x")
        tk.Label(instructions, text="3. Stop: 'F10' (or 'q'). Pause: 'F9'.", bg=bg, fg="#f38ba8", font=("Segoe UI", 9, "bold"), anchor="w").pack(fill="x")
        
        # Button Set Frame
        set_frame = tk.Frame(self, bg=bg)
        set_frame.pack(padx=24, pady=(12, 0), fill="x")
        
        self.set_btn = ttk.Button(set_frame, text="Set Next Button Location", command=self._on_set_button)
        self.set_btn.pack(side="left", fill="x", expand=True, ipady=4)
        
        self.coord_label = ttk.Label(set_frame, text="Not set", style="TLabel")
        self.coord_label.pack(side="left", padx=10)

        # Form
        form = tk.Frame(self, bg=bg)
        form.pack(padx=24, pady=(12, 0), fill="x")
        
        ttk.Label(form, text="Total Pages:", style="TLabel").grid(row=0, column=0, sticky="w", pady=4)
        self.pages_entry = ttk.Entry(form, width=20, style="TEntry")
        self.pages_entry.grid(row=0, column=1, padx=(8, 0), pady=4, sticky="w")
        ttk.Label(form, text="(leave blank for unlimited)", style="TLabel").grid(row=0, column=2, padx=(8, 0))
        
        ttk.Label(form, text="Delay (ms):", style="TLabel").grid(row=1, column=0, sticky="w", pady=4)
        self.delay_entry = ttk.Entry(form, width=20, style="TEntry")
        self.delay_entry.insert(0, "500")
        self.delay_entry.grid(row=1, column=1, padx=(8, 0), pady=4, sticky="w")
        
        # Buttons
        btn_frame = tk.Frame(self, bg=bg)
        btn_frame.pack(pady=12)
        
        self.start_btn = ttk.Button(btn_frame, text="Start Capture", style="Accent.TButton", command=self._on_start)
        self.start_btn.pack(side="left", padx=6, ipadx=12, ipady=4)
        
        self.stop_btn = ttk.Button(btn_frame, text="Stop (Hold 'q')", style="Stop.TButton", command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=6, ipadx=12, ipady=4)
        
        # Progress
        self.progress = ttk.Progressbar(self, length=500, mode="determinate", style="green.Horizontal.TProgressbar")
        self.progress.pack(padx=24, pady=(4, 6))
        
        # Log
        self.log_area = scrolledtext.ScrolledText(
            self, height=10, bg="#181825", fg="#a6adc8",
            font=("Consolas", 9), insertbackground=fg,
            relief="flat", borderwidth=0, state="disabled")
        self.log_area.pack(padx=24, pady=(0, 16), fill="both", expand=True)
    
    def _log(self, msg):
        def _append():
            self.log_area.configure(state="normal")
            self.log_area.insert("end", msg + "\n")
            self.log_area.see("end")
            self.log_area.configure(state="disabled")
        self.after(0, _append)
    
    def _set_progress(self, value):
        def _update():
            if value < 0:
                self.progress.configure(mode="indeterminate")
                self.progress.start(15)
            else:
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self.progress["value"] = value
        self.after(0, _update)
    
    def _on_done(self):
        def _finish():
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.progress.stop()
            self.progress["value"] = 100
        self.after(0, _finish)
        
    def _on_set_button(self):
        messagebox.showinfo("Set Next Button", 
                            "Instructions:\n"
                            "1. Move your mouse over the '>' (Next Page) button in VitalSource.\n"
                            "2. Press the 'n' key on your keyboard to save the location.")
        
        self.set_btn.configure(text="Waiting for 'n' key...", state="disabled")
        self.update()
        
        # Wait for key press
        while True:
            if keyboard.is_pressed('n'):
                mx, my = pyautogui.position()
                
                # Get the window to calculate relative offset
                try:
                    import pygetwindow as gw
                    # We try to find the window right now
                    win = find_vitalsource_window()
                    
                    if win:
                        off_x = mx - win.left
                        off_y = my - win.top
                        GlobalState.relative_offset = (off_x, off_y)
                        self.coord_label.configure(text=f"Offset: ({off_x}, {off_y})", foreground="#a6e3a1")
                        self.set_btn.configure(text="Set Next Button Location", state="normal")
                        messagebox.showinfo("Success", f"Location set relative to window.\nOffset: ({off_x}, {off_y})")
                    else:
                        messagebox.showerror("Error", "VitalSource window not found! Is it open?")
                        self.set_btn.configure(text="Set Next Button Location", state="normal")
                        
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to calculate offset: {e}")
                    self.set_btn.configure(text="Set Next Button Location", state="normal")

                time.sleep(0.5) # Debounce
                break
            self.update()
            time.sleep(0.05)
    
    def _on_start(self):
        if not GlobalState.relative_offset:
            messagebox.showwarning("Missing Setup", "Please set the Next Button location first!")
            return

        pages_text = self.pages_entry.get().strip()
        delay_text = self.delay_entry.get().strip()
        
        total_pages = None
        if pages_text:
            if pages_text.isdigit() and int(pages_text) > 0:
                total_pages = int(pages_text)
            else:
                messagebox.showwarning("Invalid", "Pages must be a positive number")
                return
        
        delay_ms = 500
        if delay_text:
            if delay_text.isdigit() and int(delay_text) >= 100:
                delay_ms = int(delay_text)
            else:
                messagebox.showwarning("Invalid", "Delay must be at least 100ms")
                return
        
        # Clear log
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.progress["value"] = 0
        
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.stop_event.clear()
        
        self.worker_thread = threading.Thread(
            target=run_capture,
            args=(total_pages, delay_ms, self._log, self._set_progress, self._on_done, self.stop_event),
            daemon=True
        )
        self.worker_thread.start()
    
    def _on_stop(self):
        self.stop_event.set()
        self._log("Stopping...")
        self.stop_btn.configure(state="disabled")
    
    def _on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if messagebox.askyesno("Confirm", "Capture in progress. Stop and exit?"):
                self.stop_event.set()
                self.worker_thread.join(timeout=3)
                self.destroy()
        else:
            self.destroy()


if __name__ == "__main__":
    # Check dependencies
    try:
        import pygetwindow
        import keyboard
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing Dependency",
            "Required libraries missing.\n\nRun: pip install pygetwindow pyautogui keyboard"
        )
        sys.exit(1)
    
    app = App()
    app.mainloop()
