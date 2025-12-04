"""
Excel Data Cleaner - Automated cleaning tool for Excel files
Removes rows based on specific patterns in designated columns
"""

import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path


class ExcelCleaner:
    """Handles Excel file cleaning operations"""
    
    # Column mappings (Excel column letters to 0-based indices)
    COLUMNS = {
        'Order': 'H',           # Column H (index 7)
        'Buyer PO Number': 'I', # Column I (index 8)
        'Comment': 'BO',        # Column BO (index 66)
        'ShipmentID': 'BV'      # Column BV (index 73)
    }
    
    def __init__(self, input_file):
        self.input_file = Path(input_file)
        self.df = None
        self.rows_removed = 0
        self.original_row_count = 0
        
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
            self.df = pd.read_excel(self.input_file, engine='openpyxl')
            self.original_row_count = len(self.df)
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
        if shipment_idx < len(self.df.columns):
            shipment_col = self.df.iloc[:, shipment_idx]
            keep_mask &= ~shipment_col.apply(lambda x: self.contains_pattern(x, ['FOC']))
        
        # Rule 2: Order (Column H) - Remove rows containing test, testing, M88, GB Test, GB Testing, GB
        if order_idx < len(self.df.columns):
            order_col = self.df.iloc[:, order_idx]
            order_patterns = ['test', 'testing', 'M88', 'GB Test', 'GB Testing', 'GB']
            keep_mask &= ~order_col.apply(lambda x: self.contains_pattern(x, order_patterns))
        
        # Rule 3: Buyer PO Number (Column I) - Remove rows containing test, testing, FOC
        if buyer_po_idx < len(self.df.columns):
            buyer_po_col = self.df.iloc[:, buyer_po_idx]
            buyer_patterns = ['test', 'testing', 'FOC']
            keep_mask &= ~buyer_po_col.apply(lambda x: self.contains_pattern(x, buyer_patterns))
        
        # Rule 4: Comment (Column BO) - Remove rows containing FOC, M88
        if comment_idx < len(self.df.columns):
            comment_col = self.df.iloc[:, comment_idx]
            comment_patterns = ['FOC', 'M88']
            keep_mask &= ~comment_col.apply(lambda x: self.contains_pattern(x, comment_patterns))
        
        # Apply the mask to keep only valid rows
        self.df = self.df[keep_mask]
        
        self.rows_removed = initial_count - len(self.df)
        
    def save_cleaned_file(self):
        """Save cleaned data to a new Excel file"""
        # Generate output filename
        output_filename = self.input_file.stem + "_CLEANED.xlsx"
        output_path = self.input_file.parent / output_filename
        
        try:
            self.df.to_excel(output_path, index=False, engine='openpyxl')
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


def select_file():
    """Open file picker dialog to select Excel file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File to Clean",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("All files", "*.*")
        ]
    )
    
    return file_path


def main():
    """Main entry point for the application"""
    input_file = None
    
    # Check if file was provided via drag-and-drop or command line
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # No file provided, open file picker
        input_file = select_file()
    
    if not input_file:
        messagebox.showinfo("Cancelled", "No file selected. Exiting.")
        return
    
    # Verify file exists and is an Excel file
    if not os.path.exists(input_file):
        messagebox.showerror("Error", f"File does not exist:\n{input_file}")
        return
    
    if not input_file.lower().endswith('.xlsx'):
        messagebox.showerror("Error", "Please select a valid Excel file (.xlsx)")
        return
    
    # Process the file
    cleaner = ExcelCleaner(input_file)
    output_path = cleaner.process()
    
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


if __name__ == "__main__":
    main()
