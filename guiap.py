import os
import time
import subprocess
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import sys

def get_startup_folder():
    return os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')

def create_shortcut(target, shortcut_path):
    vbs_script = f'''
    Set oWS = WScript.CreateObject("WScript.Shell")
    sLinkFile = "{shortcut_path}"
    Set oLink = oWS.CreateShortcut(sLinkFile)
    oLink.TargetPath = "{target}"
    oLink.Save
    '''
    vbs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "create_shortcut.vbs")
    with open(vbs_path, "w") as f:
        f.write(vbs_script)
    subprocess.run(["cscript", vbs_path])
    os.remove(vbs_path)

class AutoPrintGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AutoPrint")
        self.geometry("600x400")

        self.watch_thread = None
        self.watching = False
        self.watch_directory_path = os.path.dirname(os.path.abspath(__file__))

        self.create_widgets()
        self.update_watch_dir_label()

    def create_widgets(self):
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Top frame for controls
        self.top_frame = tk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X)

        self.toggle_button = tk.Button(self.top_frame, text="Start Watching", command=self.toggle_watching)
        self.toggle_button.pack(side=tk.LEFT, padx=(0, 10))

        self.select_dir_button = tk.Button(self.top_frame, text="Select Directory", command=self.select_directory)
        self.select_dir_button.pack(side=tk.LEFT)

        self.startup_var = tk.BooleanVar()
        self.startup_check = tk.Checkbutton(self.top_frame, text="Run at Startup", var=self.startup_var, command=self.toggle_startup)
        self.startup_check.pack(side=tk.RIGHT)

        self.watch_dir_label = tk.Label(self.main_frame, text="")
        self.watch_dir_label.pack(fill=tk.X, pady=5)

        self.log_area = scrolledtext.ScrolledText(self.main_frame, wrap=tk.WORD, height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def update_watch_dir_label(self):
        self.watch_dir_label.config(text=f"Watching: {self.watch_directory_path}")

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.watch_directory_path = directory
            self.update_watch_dir_label()
            if self.watching:
                self.stop_watching()
                self.start_watching()

    def toggle_watching(self):
        if self.watching:
            self.stop_watching()
        else:
            self.start_watching()

    def start_watching(self):
        self.watching = True
        self.toggle_button.config(text="Stop Watching")
        self.log_message(f"Starting to watch {self.watch_directory_path}")
        self.watch_thread = threading.Thread(target=self.watch_directory, args=(self.watch_directory_path,), daemon=True)
        self.watch_thread.start()

    def stop_watching(self):
        self.watching = False
        self.toggle_button.config(text="Start Watching")
        self.log_message("Stopped watching.")

    def toggle_startup(self):
        startup_folder = get_startup_folder()
        shortcut_path = os.path.join(startup_folder, "AutoPrint.lnk")

        if self.startup_var.get():
            target_path = sys.executable
            create_shortcut(target_path, shortcut_path)
            self.log_message("Added to startup.")
        else:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                self.log_message("Removed from startup.")

    def log_message(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def print_pdf_silently(self, pdf_path, printer_name=None):
        sumatra_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SumatraPDF", "SumatraPDF.exe")
        if not os.path.exists(pdf_path):
            self.log_message(f"Error: PDF file not found at '{pdf_path}'")
            return

        command = [sumatra_path, '-exit-on-print']

        if printer_name:
            command.extend(['-print-to', printer_name])
        else:
            command.append('-print-to-default')

        command.extend(['-silent', pdf_path])

        try:
            subprocess.run(command, check=True, creationflags=subprocess.SW_HIDE)
            self.log_message(f"Successfully sent '{pdf_path}' to the printer.")
        except FileNotFoundError:
            self.log_message(f"Error: SumatraPDF.exe not found at '{sumatra_path}'")
            self.log_message("Please ensure the path is correct.")
        except subprocess.CalledProcessError as e:
            self.log_message(f"Error during printing: {e}")

    def convert_image_to_pdf(self, image_path, pdf_path):
        try:
            img = Image.open(image_path)
            img_w, img_h = img.size

            A4_WIDTH = 827
            A4_HEIGHT = 1169
            pdf_size = (A4_WIDTH, A4_HEIGHT)

            w_ratio = A4_WIDTH / img_w
            h_ratio = A4_HEIGHT / img_h
            scale_factor = min(w_ratio, h_ratio)

            new_w = int(img_w * scale_factor)
            new_h = int(img_h * scale_factor)

            img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

            background = Image.new('RGB', pdf_size, (255, 255, 255))

            paste_x = (A4_WIDTH - new_w) // 2
            paste_y = (A4_HEIGHT - new_h) // 2

            if img_resized.mode == 'RGBA':
                background.paste(img_resized, (paste_x, paste_y), mask=img_resized.split()[3])
            else:
                background.paste(img_resized, (paste_x, paste_y))

            background.save(pdf_path, 'PDF', resolution=100.0)

            self.log_message(f"Successfully converted '{image_path}' to '{pdf_path}'")
            return True
        except Exception as e:
            self.log_message(f"Error converting image to PDF: {e}")
            return False

    def watch_directory(self, directory):
        processed_files = {os.path.join(directory, f) for f in os.listdir(directory)}
        self.log_message(f"Ignoring {len(processed_files)} existing file(s).")

        while self.watching:
            try:
                for filename in os.listdir(directory):
                    file_path = os.path.join(directory, filename)
                    if file_path not in processed_files:
                        initial_size = -1
                        while True:
                            try:
                                current_size = os.path.getsize(file_path)
                                if current_size == initial_size:
                                    break
                                initial_size = current_size
                                time.sleep(1)
                            except OSError:
                                time.sleep(1)
                                continue

                        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                            self.log_message(f"New image detected: {filename}")
                            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                            if self.convert_image_to_pdf(file_path, pdf_path):
                                processed_files.add(file_path)
                                self.print_pdf_silently(pdf_path)
                                processed_files.add(pdf_path)
                        
                        elif filename.lower().endswith(".pdf"):
                            self.log_message(f"New PDF detected: {filename}")
                            self.print_pdf_silently(file_path)
                            processed_files.add(file_path)

                time.sleep(5)
            except Exception as e:
                self.log_message(f"An error occurred: {e}")
                time.sleep(10)

if __name__ == '__main__':
    app = AutoPrintGUI()
    app.mainloop()