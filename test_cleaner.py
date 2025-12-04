"""
Test script for Excel Cleaner
Creates a sample Excel file and tests the cleaning logic
"""

import pandas as pd
from pathlib import Path
import sys

def create_test_excel():
    """Create a test Excel file with sample data"""
    
    # Create enough columns to reach column BV (73 columns)
    num_columns = 74  # 0-73 indices, BV is index 73
    
    # Initialize empty data
    data = {}
    for i in range(num_columns):
        data[f'Col_{i}'] = [''] * 20  # 20 rows
    
    df = pd.DataFrame(data)
    
    # Column H (index 7) - Order
    df.iloc[0, 7] = "Normal Order"
    df.iloc[1, 7] = "Test Order"  # Should be removed
    df.iloc[2, 7] = "M88-12345"  # Should be removed
    df.iloc[3, 7] = "GB Testing Order"  # Should be removed
    df.iloc[4, 7] = "GB Order"  # Should be removed
    df.iloc[5, 7] = "Valid Order"
    
    # Column I (index 8) - Buyer PO Number
    df.iloc[6, 8] = "PO-Testing-123"  # Should be removed
    df.iloc[7, 8] = "FOC-PO"  # Should be removed
    df.iloc[8, 8] = "Valid PO"
    
    # Column BO (index 66) - Comment
    df.iloc[9, 66] = "FOC Comment"  # Should be removed
    df.iloc[10, 66] = "M88 in comment"  # Should be removed
    df.iloc[11, 66] = "Normal comment"
    
    # Column BV (index 73) - ShipmentID
    df.iloc[12, 73] = "FOC123"  # Should be removed
    df.iloc[13, 73] = "foc-shipment"  # Should be removed (case insensitive)
    df.iloc[14, 73] = "Valid Shipment"
    
    # Add some valid rows
    df.iloc[15, 7] = "Order 15"
    df.iloc[15, 8] = "PO 15"
    df.iloc[15, 66] = "Comment 15"
    df.iloc[15, 73] = "Ship 15"
    
    df.iloc[16, 7] = "Order 16"
    df.iloc[16, 8] = "PO 16"
    df.iloc[16, 66] = "Comment 16"
    df.iloc[16, 73] = "Ship 16"
    
    # Save to Excel
    output_path = Path("test_sample.xlsx")
    df.to_excel(output_path, index=False, engine='openpyxl')
    
    print(f"✓ Created test file: {output_path}")
    print(f"  Total rows: {len(df)}")
    print(f"  Expected removals:")
    print(f"    - Row 1: 'Test Order' in column H")
    print(f"    - Row 2: 'M88-12345' in column H")
    print(f"    - Row 3: 'GB Testing Order' in column H")
    print(f"    - Row 4: 'GB Order' in column H")
    print(f"    - Row 6: 'PO-Testing-123' in column I")
    print(f"    - Row 7: 'FOC-PO' in column I")
    print(f"    - Row 9: 'FOC Comment' in column BO")
    print(f"    - Row 10: 'M88 in comment' in column BO")
    print(f"    - Row 12: 'FOC123' in column BV")
    print(f"    - Row 13: 'foc-shipment' in column BV")
    print(f"  Expected remaining: ~10 rows")
    
    return output_path

def test_cleaning_logic():
    """Test the cleaning logic without GUI"""
    from excel_cleaner import ExcelCleaner
    
    # Create test file
    test_file = create_test_excel()
    
    print("\n" + "="*60)
    print("Testing cleaning logic...")
    print("="*60)
    
    # Create cleaner instance
    cleaner = ExcelCleaner(test_file)
    
    # Load file
    if not cleaner.load_file():
        print("✗ Failed to load file")
        return False
    
    print(f"✓ Loaded file successfully")
    print(f"  Original rows: {cleaner.original_row_count}")
    
    # Validate columns
    if not cleaner.validate_columns():
        print("✗ Column validation failed")
        return False
    
    print(f"✓ All required columns present")
    
    # Clean data
    cleaner.clean_data()
    print(f"✓ Cleaning completed")
    print(f"  Rows removed: {cleaner.rows_removed}")
    print(f"  Remaining rows: {len(cleaner.df)}")
    
    # Save cleaned file
    output_path = cleaner.save_cleaned_file()
    if output_path:
        print(f"✓ Saved cleaned file: {output_path}")
    else:
        print("✗ Failed to save cleaned file")
        return False
    
    print("\n" + "="*60)
    print("TEST COMPLETED SUCCESSFULLY!")
    print("="*60)
    print(f"\nYou can now:")
    print(f"1. Open '{test_file}' to see the original data")
    print(f"2. Open '{output_path}' to see the cleaned data")
    print(f"3. Compare them to verify the cleaning rules worked correctly")
    
    return True

if __name__ == "__main__":
    try:
        test_cleaning_logic()
    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
