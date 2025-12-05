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


class ExcelCleaner:
    """Handles Excel file cleaning operations"""
    
    # Column mappings (Excel column letters to 0-based indices)
    COLUMNS = {
        'Order': 'H',           # Column H (index 7)
        'Buyer PO Number': 'I', # Column I (index 8)
        'Comment': 'BO',        # Column BO (index 66)
        'ShipmentID': 'BV'      # Column BV (index 73)
    }
    
    def __init__(self, input_file, progress_callback=None):
        self.input_file = Path(input_file)
        self.df = None
        self.rows_removed = 0
        self.original_row_count = 0
        self.progress_callback = progress_callback
    
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
    """Progress window to show cleaning status"""
    
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Processing...")
        self.window.geometry("400x200")
        self.window.resizable(False, False)
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.window.winfo_screenheight() // 2) - (200 // 2)
        self.window.geometry(f"400x200+{x}+{y}")
        
        # Make it modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Progress label
        self.label = tk.Label(
            self.window,
            text="Initializing...",
            font=("Arial", 10),
            wraplength=350
        )
        self.label.pack(pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            self.window,
            mode='indeterminate',
            length=350
        )
        self.progress.pack(pady=10)
        self.progress.start(10)
        
        # Status text
        self.status = tk.Label(
            self.window,
            text="",
            font=("Arial", 8),
            fg="gray"
        )
        self.status.pack(pady=10)
    
    def update_message(self, message):
        """Update the progress message"""
        self.label.config(text=message)
        self.window.update()
    
    def close(self):
        """Close the progress window"""
        self.progress.stop()
        self.window.destroy()


class ExcelCleanerGUI:
    """Main GUI application for Excel Cleaner"""
    
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Excel Cleaner")
        self.root.geometry("500x350")
        self.root.resizable(False, False)
        
        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.root.winfo_screenheight() // 2) - (350 // 2)
        self.root.geometry(f"500x350+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Title
        title = tk.Label(
            self.root,
            text="Excel Data Cleaner",
            font=("Arial", 18, "bold"),
            fg="#2c3e50"
        )
        title.pack(pady=20)
        
        # Subtitle
        subtitle = tk.Label(
            self.root,
            text="Remove test data and FOC entries from Excel files",
            font=("Arial", 10),
            fg="#7f8c8d"
        )
        subtitle.pack()
        
        # Drop zone frame
        self.drop_frame = tk.Frame(
            self.root,
            bg="#ecf0f1",
            relief=tk.RIDGE,
            borderwidth=2
        )
        self.drop_frame.pack(pady=30, padx=40, fill=tk.BOTH, expand=True)
        
        # Drop zone label
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📁\n\nDrag & Drop Excel File Here\n\nor",
            font=("Arial", 12),
            bg="#ecf0f1",
            fg="#34495e"
        )
        self.drop_label.pack(expand=True)
        
        # Browse button
        browse_btn = tk.Button(
            self.drop_frame,
            text="Browse Files",
            command=self.browse_file,
            font=("Arial", 10),
            bg="#3498db",
            fg="white",
            padx=20,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2"
        )
        browse_btn.pack(pady=10)
        
        # Enable drag and drop
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        
        # Info label
        info = tk.Label(
            self.root,
            text="Supported format: .xlsx files only",
            font=("Arial", 8),
            fg="#95a5a6"
        )
        info.pack(pady=5)
    
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
                    progress_callback=lambda msg: progress_window.update_message(msg)
                )
                output_path = cleaner.process()
                
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
                    self.root.after(0, lambda: messagebox.showinfo("Success", message))
                else:
                    self.root.after(0, lambda: messagebox.showerror("Error", "Cleaning process failed."))
            except Exception as e:
                self.root.after(0, progress_window.close)
                self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred:\n{str(e)}"))
        
        thread = threading.Thread(target=process_thread, daemon=True)
        thread.start()
    
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
