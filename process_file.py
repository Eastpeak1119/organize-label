import pandas as pd

import os

def main():
    # Ask for filename input (without extension)
    base_name = input("Please enter the filename (without extension): ").strip()
    filename = f"{base_name}.xlsx"
    
    if not os.path.exists(filename):
        print(f"Error: File '{filename}' not found.")
        return

    try:
        # 1. Read the file without header to extract Customer and PO
        # Assuming Customer is in Row 0 (Messrs:) and Value in next col
        # Assuming PO is in Row 1 (S/C#) and Value in next col
        # We read first few rows
        df_meta = pd.read_excel(filename, header=None, nrows=5)
        
        customer = ""
        po = ""
        
        # Simple search for metadata in first few rows
        # Based on inspection:
        # Row 0: Col 1 = "Messrs:", Col 2 = Customer
        # Row 1: Col 1 = "S/C#", Col 2 = PO
        # We'll search dynamically in the first few rows/cols just in case
        
        for r in range(len(df_meta)):
            for c in range(len(df_meta.columns) - 1):
                val = str(df_meta.iat[r, c])
                if "Messrs" in val:
                     customer = str(df_meta.iat[r, c+1]).strip()
                if "S/C" in val:
                     po = str(df_meta.iat[r, c+1]).strip()

        # Clean up customer string if it contains newlines
        customer = customer.replace('\n', ' ').strip()
        
        # 2. Find header row for data
        header_row_index = None
        for i, val in enumerate(df_meta.iloc[:, 0]):
            if 'CTN' in str(val):
                header_row_index = i
                break
        
        if header_row_index is None:
            # Fallback search in entire file if not in first 5 rows (unlikely but safe)
            df_temp = pd.read_excel(filename, header=None)
            for i, val in enumerate(df_temp.iloc[:, 0]):
                if 'CTN' in str(val):
                    header_row_index = i
                    break

        if header_row_index is None:
             print("Error: Could not find 'CTN' header row.")
             return

        # 3. Read data
        df = pd.read_excel(filename, header=header_row_index)
        
        # Ensure required columns exist
        required_cols = ['CTN', 'SKU', 'Quantity', 'N.W', 'G.W']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
             print(f"Error: Missing columns {missing_cols}")
             return

        # 4. Process Data
        # Filter out rows where CTN is NaN (end of table usually)
        df = df.dropna(subset=['CTN'])
        
        # Ensure types for aggregation
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        df['N.W'] = pd.to_numeric(df['N.W'], errors='coerce').fillna(0)
        df['G.W'] = pd.to_numeric(df['G.W'], errors='coerce').fillna(0)
        
        # Create 'Part_Entry' for aggregation: SKU * Quantity
        # Note: Quantity needs to be int if possible for the string output
        def format_part(row):
            qty = row['Quantity']
            if qty.is_integer():
                qty = int(qty)
            return f"{row['SKU']}*{qty}"

        df['Part_Entry'] = df.apply(format_part, axis=1)

        # Group by CTN
        # Aggregation rules:
        # Part: join with " / "
        # N.W: sum
        # G.W: sum
        # Quantity: sum
        
        df_grouped = df.groupby('CTN').agg({
            'Part_Entry': lambda x: ' / '.join(x),
            'N.W': 'sum',
            'G.W': 'sum',
            'Quantity': 'sum'
        }).reset_index()

        # Rename columns to match output requirement
        df_grouped = df_grouped.rename(columns={
            'Part_Entry': 'Part',
            'N.W': 'NW',
            'G.W': 'GW',
            'Quantity': 'QTY'
        })
        
        # Add Metadata
        df_grouped['Customer'] = customer
        df_grouped['PO'] = po
        
        # Reorder columns: CTN, Part, NW, GW, QTY, Customer, PO
        cols_order = ['CTN', 'Part', 'NW', 'GW', 'QTY', 'Customer', 'PO']
        df_grouped = df_grouped[cols_order]
        
        # Sort by CTN (numeric sort if possible)
        # CTN might be read as string or float/int. Try to convert to numeric for sorting
        df_grouped['CTN_Numeric'] = pd.to_numeric(df_grouped['CTN'], errors='coerce')
        df_grouped = df_grouped.sort_values('CTN_Numeric').drop(columns=['CTN_Numeric'])

        # 5. Output
        output_filename = f"{filename}-res.csv"
        df_grouped.to_csv(output_filename, index=False)
        print(f"Successfully created '{output_filename}'")
        
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
