import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk # Import ttk
import threading
import os
import sys

# Import the processing logic
try:
    import calculate_distances
except ImportError:
    # If running from a different directory, add the directory to path
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import calculate_distances

# Try to import tkinterdnd2 for Drag & Drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
    BaseClass = TkinterDnD.Tk
except ImportError:
    HAS_DND = False
    BaseClass = tk.Tk
    print("Warning: tkinterdnd2 not found. Drag & Drop will not be available.")

class DistanceCalculatorGUI(BaseClass):
    def __init__(self):
        super().__init__()
        
        self.title("Astronomical Distance Calculator")
        self.geometry("600x500")
        
        # Configure grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1) # Log area expands
        
        # --- Input Section ---
        input_frame = ttk.LabelFrame(self, text="Input PI AnnotateImage objects.txt File", padding=10)
        input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="Input File:").grid(row=0, column=0, sticky="w")
        self.input_entry = ttk.Entry(input_frame)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=5)
        
        browse_btn = ttk.Button(input_frame, text="Browse...", command=self.browse_file)
        browse_btn.grid(row=0, column=2)
        
        # Drag & Drop Label
        if HAS_DND:
            dnd_label = tk.Label(input_frame, text="**OR** Drag & Drop PI AnnotateImage objects.txt File Here", relief="sunken", bg="#f0f0f0", fg="#9b9aab", height=3)
            dnd_label.grid(row=1, column=0, columnspan=3, sticky="ew", pady=10)
            
            # Register Drop Target
            dnd_label.drop_target_register(DND_FILES)
            dnd_label.dnd_bind('<<Drop>>', self.drop_file)
            
            # Also register the main window
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self.drop_file)
        else:
            ttk.Label(input_frame, text="(Install tkinterdnd2 for Drag & Drop support)", foreground="gray").grid(row=1, column=0, columnspan=3, pady=5)

        # --- Output Section ---
        ttk.Label(input_frame, text="Output Dir:").grid(row=2, column=0, sticky="w", pady=5)
        self.output_entry = ttk.Entry(input_frame)
        self.output_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        output_browse_btn = ttk.Button(input_frame, text="Browse...", command=self.browse_output_directory)
        output_browse_btn.grid(row=2, column=2, pady=5)

        # --- Actions ---
        action_frame = ttk.Frame(self)
        action_frame.grid(row=1, column=0, pady=5)
        
        self.run_btn = ttk.Button(action_frame, text="Calculate Distances", command=self.start_processing)
        self.run_btn.pack(side="left", padx=5)
        
        ttk.Button(action_frame, text="Exit", command=self.quit).pack(side="left", padx=5)
        
        # --- Log Area ---
        log_frame = ttk.LabelFrame(self, text="Log", padding=10)
        log_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_text.pack(fill="both", expand=True)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").grid(row=3, column=0, sticky="ew")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)

    def browse_output_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, directory)

    def drop_file(self, event):
        # Event data might be enclosed in braces if it contains spaces
        data = event.data
        if data.startswith('{') and data.endswith('}'):
            data = data[1:-1]
        
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, data)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        # Force update to show log immediately
        self.update_idletasks()

    def start_processing(self):
        input_file = self.input_entry.get().strip()
        if not input_file:
            messagebox.showerror("Error", "Please select an input file.")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"File not found: {input_file}")
            return
            
        self.run_btn.config(state='disabled')
        self.status_var.set("Processing...")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # Get output directory
        output_dir = self.output_entry.get().strip()
        output_file = None
        if output_dir:
            if not os.path.exists(output_dir):
                 messagebox.showerror("Error", f"Output directory not found: {output_dir}")
                 return
            output_file = os.path.join(output_dir, 'Astronomical_Distances_TAP.xlsx')

        # Run in a separate thread
        thread = threading.Thread(target=self.run_logic, args=(input_file, output_file))
        thread.daemon = True
        thread.start()

    def run_logic(self, input_file, output_file=None):
        try:
            # We pass self.log as the callback
            # Since tkinter is not thread-safe for GUI updates, we should ideally use after() or queue
            # But for simple print-like logging, it often works or we can wrap it.
            # Let's wrap it properly.
            
            def safe_log(msg):
                self.after(0, lambda: self.log(msg))
                
            calculate_distances.process_file(input_file, output_file=output_file, progress_callback=safe_log)
            
            self.after(0, lambda: messagebox.showinfo("Success", "Processing complete!"))
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {e}"))
            self.after(0, lambda: self.log(f"ERROR: {e}"))
        finally:
            self.after(0, lambda: self.run_btn.config(state='normal'))
            self.after(0, lambda: self.status_var.set("Ready"))

if __name__ == "__main__":
    app = DistanceCalculatorGUI()
    app.mainloop()
