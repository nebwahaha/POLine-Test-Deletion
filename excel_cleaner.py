"""
Excel Data Cleaner - Automated cleaning tool for Excel files
Removes rows based on specific patterns in designated columns
"""

import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import threading
import math


class ExcelCleaner:
    """Handles Excel file cleaning operations"""
    
    # Column mappings (Excel column letters to 0-based indices)
    COLUMNS = {
        'Order': 'H',           # Column H (index 7)
        'Buyer PO Number': 'I', # Column I (index 8)
        'Comment': 'BO',        # Column BO (index 66)
        'ShipmentID': 'BV'      # Column BV (index 73)
    }
    
    def __init__(self, input_file, progress_callback=None, save_deleted=False):
        self.input_file = Path(input_file)
        self.df = None
        self.rows_removed = 0
        self.original_row_count = 0
        self.progress_callback = progress_callback
        self.save_deleted = save_deleted
        self.deleted_rows = None
    
    def update_progress(self, message):
        """Update progress message if callback is provided"""
        if self.progress_callback:
            self.progress_callback(message)
        
    def column_letter_to_index(self, col_letter):
        """Convert Excel column letter to 0-based index"""
        col_letter = col_letter.upper()
        result = 0
        for char in col_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    def load_file(self):
        """Load Excel file into pandas DataFrame"""
        try:
            self.update_progress("Loading Excel file...")
            self.df = pd.read_excel(self.input_file, engine='openpyxl')
            self.original_row_count = len(self.df)
            self.update_progress(f"Loaded {self.original_row_count} rows")
            return True
        except FileNotFoundError:
            messagebox.showerror("Error", f"File not found: {self.input_file}")
            return False
        except PermissionError:
            messagebox.showerror("Error", f"Permission denied. File may be open in another program:\n{self.input_file}")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file:\n{str(e)}")
            return False
    
    def validate_columns(self):
        """Verify that required columns exist"""
        self.update_progress("Validating columns...")
        required_indices = []
        missing_columns = []
        
        for col_name, col_letter in self.COLUMNS.items():
            col_index = self.column_letter_to_index(col_letter)
            required_indices.append(col_index)
            
            if col_index >= len(self.df.columns):
                missing_columns.append(f"{col_name} (Column {col_letter})")
        
        if missing_columns:
            messagebox.showerror(
                "Missing Columns",
                f"The following required columns are missing:\n" + "\n".join(missing_columns)
            )
            return False
        self.update_progress("Column validation complete")
        return True
    
    def contains_pattern(self, value, patterns):
        """Check if value contains any of the patterns (case-insensitive substring match)"""
        if pd.isna(value):
            return False
        
        value_str = str(value).lower()
        for pattern in patterns:
            if pattern.lower() in value_str:
                return True
        return False
    
    def clean_data(self):
        """Apply all cleaning rules and remove matching rows"""
        initial_count = len(self.df)
        
        # Get column indices
        order_idx = self.column_letter_to_index('H')
        buyer_po_idx = self.column_letter_to_index('I')
        comment_idx = self.column_letter_to_index('BO')
        shipment_idx = self.column_letter_to_index('BV')
        
        # Create a mask for rows to keep (True = keep, False = remove)
        keep_mask = pd.Series([True] * len(self.df), index=self.df.index)
        
        # Rule 1: ShipmentID (Column BV) - Remove rows containing "FOC"
        self.update_progress("Cleaning ShipmentID column...")
        if shipment_idx < len(self.df.columns):
            shipment_col = self.df.iloc[:, shipment_idx]
            keep_mask &= ~shipment_col.apply(lambda x: self.contains_pattern(x, ['FOC']))
        
        # Rule 2: Order (Column H) - Remove rows containing test, testing, M88, GB Test, GB Testing, GB
        self.update_progress("Cleaning Order column...")
        if order_idx < len(self.df.columns):
            order_col = self.df.iloc[:, order_idx]
            order_patterns = ['test', 'testing', 'M88', 'GB Test', 'GB Testing', 'GB']
            keep_mask &= ~order_col.apply(lambda x: self.contains_pattern(x, order_patterns))
        
        # Rule 3: Buyer PO Number (Column I) - Remove rows containing test, testing, FOC
        self.update_progress("Cleaning Buyer PO Number column...")
        if buyer_po_idx < len(self.df.columns):
            buyer_po_col = self.df.iloc[:, buyer_po_idx]
            buyer_patterns = ['test', 'testing', 'FOC']
            keep_mask &= ~buyer_po_col.apply(lambda x: self.contains_pattern(x, buyer_patterns))
        
        # Rule 4: Comment (Column BO) - Remove rows containing FOC, M88
        self.update_progress("Cleaning Comment column...")
        if comment_idx < len(self.df.columns):
            comment_col = self.df.iloc[:, comment_idx]
            comment_patterns = ['FOC', 'M88']
            keep_mask &= ~comment_col.apply(lambda x: self.contains_pattern(x, comment_patterns))
        
        # Store deleted rows if save_deleted is enabled
        if self.save_deleted:
            self.update_progress("Storing deleted rows...")
            self.deleted_rows = self.df[~keep_mask].copy()
        
        # Apply the mask to keep only valid rows
        self.update_progress("Applying filters...")
        self.df = self.df[keep_mask]
        
        self.rows_removed = initial_count - len(self.df)
        self.update_progress(f"Removed {self.rows_removed} rows")
        
    def save_cleaned_file(self):
        """Save cleaned data to a new Excel file"""
        # Generate output filename
        output_filename = self.input_file.stem + "_CLEANED.xlsx"
        output_path = self.input_file.parent / output_filename
        
        try:
            self.update_progress("Saving cleaned file...")
            self.df.to_excel(output_path, index=False, engine='openpyxl')
            self.update_progress("File saved successfully!")
            return output_path
        except PermissionError:
            messagebox.showerror("Error", f"Cannot write to file. It may be open in another program:\n{output_path}")
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save cleaned file:\n{str(e)}")
            return None
    
    def save_deleted_file(self):
        """Save deleted rows to a separate Excel file"""
        if self.deleted_rows is None or len(self.deleted_rows) == 0:
            return None
        
        # Generate output filename
        output_filename = self.input_file.stem + "_DELETED.xlsx"
        output_path = self.input_file.parent / output_filename
        
        try:
            self.update_progress("Saving deleted rows file...")
            self.deleted_rows.to_excel(output_path, index=False, engine='openpyxl')
            self.update_progress("Deleted rows file saved!")
            return output_path
        except PermissionError:
            messagebox.showerror("Error", f"Cannot write to file. It may be open in another program:\n{output_path}")
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save deleted rows file:\n{str(e)}")
            return None
    
    def process(self):
        """Main processing pipeline"""
        if not self.load_file():
            return None
        
        if not self.validate_columns():
            return None
        
        self.clean_data()
        output_path = self.save_cleaned_file()
        
        return output_path


class ProgressWindow:
    """Progress window to show cleaning status with circular loading animation"""
    
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Processing...")
        self.window.geometry("500x350")
        self.window.resizable(False, False)
        self.window.configure(bg='white')
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.window.winfo_screenheight() // 2) - (350 // 2)
        self.window.geometry(f"500x350+{x}+{y}")
        
        # Make it modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Canvas for circular loading animation
        self.canvas = tk.Canvas(
            self.window,
            width=80,
            height=80,
            bg='white',
            highlightthickness=0
        )
        self.canvas.pack(pady=20)
        
        # Progress label
        self.label = tk.Label(
            self.window,
            text="Initializing...",
            font=("Segoe UI", 11),
            wraplength=450,
            bg='white',
            fg='#2c3e50'
        )
        self.label.pack(pady=10)
        
        # Warning label
        warning_label = tk.Label(
            self.window,
            text="⚠ Warning! Please wait until this loading screen is gone\nbefore opening the _cleaned Excel file.",
            font=("Segoe UI", 11, "bold"),
            wraplength=450,
            bg='white',
            fg='#e74c3c',
            justify=tk.CENTER
        )
        warning_label.pack(pady=15)
        
        # Animation properties
        self.angle = 0
        self.num_bars = 12
        self.bar_length = 15
        self.bar_width = 4
        self.radius = 25
        self.animation_running = True
        
        # Start animation
        self.animate()
    
    def draw_spinner(self):
        """Draw the circular loading spinner"""
        self.canvas.delete("all")
        
        for i in range(self.num_bars):
            # Calculate angle for this bar
            bar_angle = (360 / self.num_bars) * i + self.angle
            rad = math.radians(bar_angle)
            
            # Calculate start and end points
            x1 = 40 + (self.radius - self.bar_length) * math.cos(rad)
            y1 = 40 + (self.radius - self.bar_length) * math.sin(rad)
            x2 = 40 + self.radius * math.cos(rad)
            y2 = 40 + self.radius * math.sin(rad)
            
            # Calculate opacity based on position (fade effect)
            opacity_index = (i - int(self.angle / (360 / self.num_bars))) % self.num_bars
            opacity = int(255 * (1 - opacity_index / self.num_bars))
            color = f'#{opacity:02x}{opacity:02x}{opacity:02x}'
            
            # Draw the bar
            self.canvas.create_line(
                x1, y1, x2, y2,
                width=self.bar_width,
                fill=color,
                capstyle=tk.ROUND
            )
    
    def animate(self):
        """Animate the spinner"""
        if self.animation_running:
            self.draw_spinner()
            self.angle = (self.angle + 30) % 360
            self.window.after(100, self.animate)
    
    def update_message(self, message):
        """Update the progress message"""
        self.label.config(text=message)
        self.window.update()
    
    def close(self):
        """Close the progress window"""
        self.animation_running = False
        self.window.destroy()


class RoundedButton(tk.Canvas):
    """Custom rounded button widget"""
    def __init__(self, parent, text, command, bg_color="#6366f1", hover_color="#4f46e5", 
                 fg_color="white", width=200, height=50, font_size=12, parent_bg="#0f172a", **kwargs):
        tk.Canvas.__init__(self, parent, width=width, height=height, bg=parent_bg, 
                          highlightthickness=0, relief=tk.FLAT, **kwargs)
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.fg_color = fg_color
        self.text = text
        self.font_size = font_size
        self.current_color = bg_color
        self.width = width
        self.height = height
        self.parent_bg = parent_bg
        
        self.bind("<Button-1>", lambda e: self.command() if self.command else None)
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        
        self.draw()
    
    def draw(self):
        self.delete("all")
        self.create_rounded_rectangle(2, 2, self.width-2, self.height-2, 
                                     radius=12, fill=self.current_color, outline="")
        self.create_text(self.width//2, self.height//2, text=self.text, 
                        fill=self.fg_color, font=("Segoe UI", self.font_size, "bold"))
    
    def create_rounded_rectangle(self, x1, y1, x2, y2, radius=20, **kwargs):
        points = [
            x1+radius, y1,
            x1+radius, y1,
            x2-radius, y1,
            x2-radius, y1,
            x2, y1,
            x2, y1+radius,
            x2, y1+radius,
            x2, y2-radius,
            x2, y2-radius,
            x2, y2,
            x2-radius, y2,
            x2-radius, y2,
            x1+radius, y2,
            x1+radius, y2,
            x1, y2,
            x1, y2-radius,
            x1, y2-radius,
            x1, y1+radius,
            x1, y1+radius,
            x1, y1
        ]
        return self.create_polygon(points, **kwargs, smooth=True)
    
    def on_enter(self, event):
        self.current_color = self.hover_color
        self.draw()
        self.config(cursor="hand2")
    
    def on_leave(self, event):
        self.current_color = self.bg_color
        self.draw()


class ExcelCleanerGUI:
    """Main GUI application for Excel Cleaner"""
    
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Excel Cleaner")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        self.root.configure(bg="#0f172a")
        
        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"700x600+{x}+{y}")
        
        self.current_screen = "main"
        self.save_deleted_var = tk.BooleanVar(value=False)
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Clear the root window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        self.root.configure(bg="#0f172a")
        
        # Main container
        main_container = tk.Frame(self.root, bg="#0f172a")
        main_container.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        # Title
        title = tk.Label(
            main_container,
            text="Excel Data Cleaner",
            font=("Segoe UI", 32, "bold"),
            fg="#ffffff",
            bg="#0f172a"
        )
        title.pack(pady=(0, 10))
        
        # Subtitle
        subtitle = tk.Label(
            main_container,
            text="Remove test data and FOC entries from Excel files",
            font=("Segoe UI", 12),
            fg="#94a3b8",
            bg="#0f172a"
        )
        subtitle.pack(pady=(0, 40))
        
        # Drop zone frame with gradient-like appearance
        self.drop_frame = tk.Frame(
            main_container,
            bg="#1e293b",
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=2,
            highlightbackground="#334155",
            highlightcolor="#6366f1"
        )
        self.drop_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 30))
        
        # Drop zone label
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📁",
            font=("Segoe UI", 48),
            bg="#1e293b",
            fg="#64748b"
        )
        self.drop_label.pack(expand=True, pady=(30, 10))
        
        # Drag & drop text
        drag_text = tk.Label(
            self.drop_frame,
            text="Drag & Drop Excel File Here",
            font=("Segoe UI", 14, "bold"),
            bg="#1e293b",
            fg="#e2e8f0"
        )
        drag_text.pack(pady=(0, 5))
        
        # Or text
        or_text = tk.Label(
            self.drop_frame,
            text="or",
            font=("Segoe UI", 11),
            bg="#1e293b",
            fg="#94a3b8"
        )
        or_text.pack(pady=(5, 20))
        
        # Browse button with custom rounded style
        browse_btn = RoundedButton(
            self.drop_frame,
            text="Browse Files",
            command=self.browse_file,
            bg_color="#6366f1",
            hover_color="#4f46e5",
            fg_color="white",
            width=180,
            height=48,
            font_size=12
        )
        browse_btn.pack(pady=(0, 30))
        
        # Enable drag and drop
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        
        # Checkbox frame
        checkbox_frame = tk.Frame(main_container, bg="#0f172a")
        checkbox_frame.pack(fill=tk.X, pady=(15, 0))
        
        # Create delete data checkbox
        checkbox = tk.Checkbutton(
            checkbox_frame,
            text="Create separate file for deleted data",
            variable=self.save_deleted_var,
            font=("Segoe UI", 10),
            bg="#0f172a",
            fg="#e2e8f0",
            activebackground="#0f172a",
            activeforeground="#6366f1",
            selectcolor="#0f172a",
            highlightthickness=0,
            bd=0
        )
        checkbox.pack(side=tk.LEFT, padx=5)
        
        # Bottom info frame
        bottom_frame = tk.Frame(main_container, bg="#0f172a", height=60)
        bottom_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Info label
        info = tk.Label(
            bottom_frame,
            text="✓ Supported format: .xlsx files only",
            font=("Segoe UI", 10),
            fg="#64748b",
            bg="#0f172a"
        )
        info.pack(side=tk.LEFT, pady=10)
        
        # Info button in bottom right corner
        info_btn = RoundedButton(
            self.root,
            text="ℹ",
            command=self.show_info_screen,
            bg_color="#6366f1",
            hover_color="#4f46e5",
            fg_color="white",
            width=55,
            height=55,
            font_size=20
        )
        info_btn.place(relx=0.95, rely=0.98, anchor=tk.SE)
    
    def on_drop(self, event):
        """Handle file drop event"""
        file_path = event.data
        # Remove curly braces if present (Windows drag-drop adds them)
        file_path = file_path.strip('{}')
        self.process_file(file_path)
    
    def browse_file(self):
        """Open file browser dialog"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File to Clean",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.process_file(file_path)
    
    def process_file(self, file_path):
        """Process the selected file"""
        # Validate file
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File does not exist:\n{file_path}")
            return
        
        if not file_path.lower().endswith('.xlsx'):
            messagebox.showerror("Error", "Please select a valid Excel file (.xlsx)")
            return
        
        # Create progress window
        progress_window = ProgressWindow(self.root)
        
        # Process in separate thread to keep UI responsive
        def process_thread():
            try:
                cleaner = ExcelCleaner(
                    file_path,
                    progress_callback=lambda msg: progress_window.update_message(msg),
                    save_deleted=self.save_deleted_var.get()
                )
                output_path = cleaner.process()
                
                # Save deleted file if checkbox is enabled
                deleted_path = None
                if self.save_deleted_var.get():
                    deleted_path = cleaner.save_deleted_file()
                
                # Close progress window
                self.root.after(0, progress_window.close)
                
                if output_path:
                    message = (
                        f"✓ Cleaning completed successfully!\n\n"
                        f"Original rows: {cleaner.original_row_count}\n"
                        f"Rows removed: {cleaner.rows_removed}\n"
                        f"Remaining rows: {len(cleaner.df)}\n\n"
                        f"Cleaned file saved to:\n{output_path}"
                    )
                    
                    if deleted_path:
                        message += f"\n\nDeleted rows file saved to:\n{deleted_path}"
                    
                    self.root.after(0, lambda: messagebox.showinfo("Success", message))
                else:
                    self.root.after(0, lambda: messagebox.showerror("Error", "Cleaning process failed."))
            except Exception as e:
                self.root.after(0, progress_window.close)
                self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred:\n{str(e)}"))
        
        thread = threading.Thread(target=process_thread, daemon=True)
        thread.start()
    
    def show_info_screen(self):
        """Display the information screen with data cleaning rules"""
        self.current_screen = "info"
        
        # Clear the root window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        self.root.configure(bg="#0f172a")
        
        # Header frame
        header_frame = tk.Frame(self.root, bg="#1e293b", relief=tk.FLAT, borderwidth=0)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        
        # Back button in top left with improved styling
        back_btn = tk.Button(
            header_frame,
            text="← Back",
            command=self.back_to_main,
            font=("Segoe UI", 10, "bold"),
            bg="#6366f1",
            fg="white",
            padx=20,
            pady=12,
            relief=tk.FLAT,
            cursor="hand2",
            activebackground="#4f46e5",
            activeforeground="white",
            bd=0,
            highlightthickness=0
        )
        back_btn.pack(anchor=tk.NW, padx=15, pady=12)
        
        # Bind hover effects to back button
        back_btn.bind("<Enter>", lambda e: back_btn.config(bg="#4f46e5"))
        back_btn.bind("<Leave>", lambda e: back_btn.config(bg="#6366f1"))
        
        # Title
        title = tk.Label(
            self.root,
            text="Data Cleaning Rules",
            font=("Segoe UI", 24, "bold"),
            fg="#ffffff",
            bg="#0f172a"
        )
        title.pack(pady=20)
        
        # Create a scrollable frame for the information
        canvas = tk.Canvas(self.root, bg="#0f172a", highlightthickness=0, relief=tk.FLAT)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#0f172a")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _on_mousewheel_linux(event):
            if event.num == 5:
                canvas.yview_scroll(3, "units")
            elif event.num == 4:
                canvas.yview_scroll(-3, "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel_linux)
        canvas.bind_all("<Button-5>", _on_mousewheel_linux)
        
        # Information content
        info_text = """The tool removes entire rows based on substring matches (case-insensitive) in specific columns:

1. ShipmentID (Column BV)
   • Removes rows containing: "FOC"

2. Order (Column H)
   • Removes rows containing any of:
     - "test" or "testing"
     - "M88" (including variants like "M880123")
     - "GB Test", "GB Testing", or "GB" alone

3. Buyer PO Number (Column I)
   • Removes rows containing:
     - "test" or "testing"
     - "FOC"

4. Comment (Column BO)
   • Removes rows containing:
     - "FOC"
     - "M88"

Output:
   • Creates a new file: <original_filename>_CLEANED.xlsx
   • Saved in the same directory as the input file
   • Displays statistics: original rows, rows removed, remaining rows

Note: All matching is case-insensitive and works on substrings.
Example: "M880123" will match "M88" and be removed."""
        
        info_label = tk.Label(
            scrollable_frame,
            text=info_text,
            font=("Segoe UI", 10),
            fg="#e2e8f0",
            justify=tk.LEFT,
            wraplength=600,
            bg="#0f172a"
        )
        info_label.pack(padx=30, pady=20, anchor="w")
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True, padx=0, pady=10)
        scrollbar.pack(side="right", fill="y", padx=5)
    
    def back_to_main(self):
        """Return to the main screen"""
        self.current_screen = "main"
        self.setup_ui()
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()


def main():
    """Main entry point for the application"""
    # Check if file was provided via command line (for backward compatibility)
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        
        # Verify file exists and is an Excel file
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"File does not exist:\n{input_file}")
            return
        
        if not input_file.lower().endswith('.xlsx'):
            messagebox.showerror("Error", "Please select a valid Excel file (.xlsx)")
            return
        
        # Create a temporary root for progress window
        root = tk.Tk()
        root.withdraw()
        
        # Create progress window
        progress_window = ProgressWindow(root)
        
        # Process the file
        cleaner = ExcelCleaner(
            input_file,
            progress_callback=lambda msg: progress_window.update_message(msg)
        )
        output_path = cleaner.process()
        
        progress_window.close()
        
        if output_path:
            message = (
                f"✓ Cleaning completed successfully!\n\n"
                f"Original rows: {cleaner.original_row_count}\n"
                f"Rows removed: {cleaner.rows_removed}\n"
                f"Remaining rows: {len(cleaner.df)}\n\n"
                f"Cleaned file saved to:\n{output_path}"
            )
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", "Cleaning process failed.")
        
        root.destroy()
    else:
        # No file provided, launch GUI
        app = ExcelCleanerGUI()
        app.run()


if __name__ == "__main__":
    main()
