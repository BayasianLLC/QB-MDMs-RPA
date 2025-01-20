import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import subprocess
import sys
import os
from datetime import datetime
import signal

class StreamRedirector:
    def __init__(self, queue):
        self.queue = queue

    def write(self, text):
        self.queue.put(text)

    def flush(self):
        pass

class MDMProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MDM Processor Control Panel")
        self.root.minsize(800, 600)
        
        # Get the directory where the UI script is located
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Configure root grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        self.processes = {}
        self.script_status = {}
        
        self.create_widgets()
        
        self.output_queue = queue.Queue()
        self.queue_thread = threading.Thread(target=self.process_queue, daemon=True)
        self.queue_thread.start()

    def create_script_controls(self, parent, script_name, script_file, row):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=5)
        frame.grid_columnconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(frame, text=script_name, font=('Arial', 10, 'bold'))
        title_label.grid(row=0, column=0, sticky="w", padx=10)

        # Status with custom styling
        status_frame = ttk.Frame(frame)
        status_frame.grid(row=0, column=1, sticky="w", padx=10)
        
        status_label = ttk.Label(status_frame, text="Status:", font=('Arial', 10))
        status_label.grid(row=0, column=0, padx=(0, 5))
        
        status_value = ttk.Label(status_frame, text="Not Running", foreground="red", font=('Arial', 10))
        status_value.grid(row=0, column=1)
        self.script_status[script_file] = status_value

        # Buttons frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=0, column=2, padx=5)

        # Start button
        start_btn = ttk.Button(
            button_frame,
            text="Start",
            command=lambda: self.start_script(script_file),
            width=10
        )
        start_btn.grid(row=0, column=0, padx=2)

        # Stop button
        stop_btn = ttk.Button(
            button_frame,
            text="Stop",
            command=lambda: self.stop_script(script_file),
            width=10
        )
        stop_btn.grid(row=0, column=1, padx=2)

        # Add separator
        separator = ttk.Separator(parent, orient='horizontal')
        separator.grid(row=row+1, column=0, sticky="ew", pady=5)

    def create_widgets(self):
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(5, weight=1)

        # Scripts section title
        scripts_label = ttk.Label(main_frame, text="MDM Processors", font=('Arial', 12, 'bold'))
        scripts_label.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Script controls
        self.create_script_controls(main_frame, "PSE MDM Processor", "pse_mdm_processor.py", 1)
        self.create_script_controls(main_frame, "PSEG MDM Processor", "pseg_mdm_processor.py", 3)
        self.create_script_controls(main_frame, "SCE MDM Processor", "sce_mdm_processor.py", 5)
        self.create_script_controls(main_frame, "SDGE MDM Processor", "sdge_mdm_processor.py", 7)

        # Console section
        console_label = ttk.Label(main_frame, text="Console Output", font=('Arial', 12, 'bold'))
        console_label.grid(row=9, column=0, sticky="w", pady=(10, 5))

        self.console = scrolledtext.ScrolledText(
            main_frame,
            wrap=tk.WORD,
            background='black',
            foreground='#00FF00',
            font=('Consolas', 10),
            height=15
        )
        self.console.grid(row=10, column=0, sticky="nsew", pady=(0, 10))

        # Button panel
        button_panel = ttk.Frame(main_frame)
        button_panel.grid(row=11, column=0, sticky="ew")
        button_panel.grid_columnconfigure(1, weight=1)

        clear_button = ttk.Button(
            button_panel,
            text="Clear Console",
            command=self.clear_console,
            width=15
        )
        clear_button.grid(row=0, column=0, padx=5)

        stop_all_button = ttk.Button(
            button_panel,
            text="Stop All Scripts",
            command=self.stop_all_scripts,
            width=15
        )
        stop_all_button.grid(row=0, column=2, padx=5)

    def start_script(self, script_file):
        script_path = os.path.join(self.script_dir, script_file)
        
        if not os.path.exists(script_path):
            error_msg = f"Script file not found: {script_path}"
            self.log_message(f"ERROR: {error_msg}")
            messagebox.showerror("Error", error_msg)
            return
            
        if script_file not in self.processes or self.processes[script_file].poll() is not None:
            try:
                process = subprocess.Popen(
                    [sys.executable, script_path],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    universal_newlines=True,
                    bufsize=1
                )
                
                self.processes[script_file] = process
                self.script_status[script_file].config(text="Running", foreground="green")
                
                threading.Thread(
                    target=self.read_output,
                    args=(process, script_file),
                    daemon=True
                ).start()
                
                self.log_message(f"Started {script_file}")
            except Exception as e:
                error_msg = f"Error starting {script_file}: {str(e)}"
                self.log_message(f"ERROR: {error_msg}")
                messagebox.showerror("Error", error_msg)
        else:
            self.log_message(f"{script_file} is already running")

    def stop_script(self, script_file):
        if script_file in self.processes and self.processes[script_file].poll() is None:
            try:
                if os.name == 'nt':  # Windows
                    self.processes[script_file].terminate()
                else:  # Unix/Linux
                    os.kill(self.processes[script_file].pid, signal.SIGTERM)
                
                self.script_status[script_file].config(text="Not Running", foreground="red")
                self.log_message(f"Stopped {script_file}")
            except Exception as e:
                error_msg = f"Error stopping {script_file}: {str(e)}"
                self.log_message(f"ERROR: {error_msg}")
                messagebox.showerror("Error", error_msg)
        else:
            self.log_message(f"{script_file} is not running")

    def stop_all_scripts(self):
        for script_file in list(self.processes.keys()):
            self.stop_script(script_file)
        self.log_message("All scripts stopped")

    def read_output(self, process, script_file):
        for line in process.stdout:
            self.log_message(f"[{script_file}] {line.strip()}")
        for line in process.stderr:
            self.log_message(f"[{script_file}] ERROR: {line.strip()}")
        
        # Update status when process ends
        self.script_status[script_file].config(text="Not Running", foreground="red")
        if script_file in self.processes:
            del self.processes[script_file]

    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.output_queue.put(f"[{timestamp}] {message}\n")

    def process_queue(self):
        while True:
            try:
                message = self.output_queue.get(timeout=0.1)
                self.console.insert(tk.END, message)
                self.console.see(tk.END)
            except queue.Empty:
                self.root.after(100, self.process_queue)
                break

    def clear_console(self):
        self.console.delete(1.0, tk.END)

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to close the application? This will stop all running scripts."):
            self.stop_all_scripts()
            self.root.destroy()

def main():
    root = tk.Tk()
    
    # Configure style
    style = ttk.Style()
    style.configure('TFrame', background='#f0f0f0')
    style.configure('TButton', padding=5)
    style.configure('TLabel', background='#f0f0f0')
    
    app = MDMProcessorUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()