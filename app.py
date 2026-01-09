import streamlit as st
import pandas as pd
import io

def process_excel(uploaded_file):
    """
    Processes the uploaded Excel file using the logic from process_file.py.
    Returns a tuple (processed_df, output_filename, error_message).
    """
    try:
        # Generate output filename
        original_filename = uploaded_file.name
        base_name = original_filename.rsplit('.', 1)[0]
        output_filename = f"{base_name}-res.csv"

        # 1. Read the file without header to extract Customer and PO
        # Need to reset pointer because we might read this multiple times
        uploaded_file.seek(0)
        df_meta = pd.read_excel(uploaded_file, header=None, nrows=5)
        
        customer = ""
        po = ""
        
        for r in range(len(df_meta)):
            for c in range(len(df_meta.columns) - 1):
                val = str(df_meta.iat[r, c])
                if "Messrs" in val:
                     customer = str(df_meta.iat[r, c+1]).strip()
                if "S/C" in val:
                     po = str(df_meta.iat[r, c+1]).strip()

        customer = customer.replace('\n', ' ').strip()
        
        # 2. Find header row for data
        header_row_index = None
        for i, val in enumerate(df_meta.iloc[:, 0]):
            if 'CTN' in str(val):
                header_row_index = i
                break
        
        if header_row_index is None:
            # Fallback search
            uploaded_file.seek(0)
            df_temp = pd.read_excel(uploaded_file, header=None)
            for i, val in enumerate(df_temp.iloc[:, 0]):
                if 'CTN' in str(val):
                    header_row_index = i
                    break

        if header_row_index is None:
             return None, None, "Error: Could not find 'CTN' header row."

        # 3. Read data
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row_index)
        
        required_cols = ['CTN', 'SKU', 'Quantity', 'N.W', 'G.W']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
             return None, None, f"Error: Missing columns {missing_cols}"

        # 4. Process Data
        df = df.dropna(subset=['CTN'])
        
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        df['N.W'] = pd.to_numeric(df['N.W'], errors='coerce').fillna(0)
        df['G.W'] = pd.to_numeric(df['G.W'], errors='coerce').fillna(0)
        
        def format_part(row):
            qty = row['Quantity']
            if qty.is_integer():
                qty = int(qty)
            return f"{row['SKU']}*{qty}"

        df['Part_Entry'] = df.apply(format_part, axis=1)

        df_grouped = df.groupby('CTN').agg({
            'Part_Entry': lambda x: ' / '.join(x),
            'N.W': 'sum',
            'G.W': 'sum',
            'Quantity': 'sum'
        }).reset_index()

        df_grouped = df_grouped.rename(columns={
            'Part_Entry': 'Part',
            'N.W': 'NW',
            'G.W': 'GW',
            'Quantity': 'QTY'
        })
        
        df_grouped['Customer'] = customer
        df_grouped['PO'] = po
        
        cols_order = ['CTN', 'Part', 'NW', 'GW', 'QTY', 'Customer', 'PO']
        df_grouped = df_grouped[cols_order]
        
        df_grouped['CTN_Numeric'] = pd.to_numeric(df_grouped['CTN'], errors='coerce')
        df_grouped = df_grouped.sort_values('CTN_Numeric').drop(columns=['CTN_Numeric'])

        return df_grouped, output_filename, None
        
    except Exception as e:
        return None, None, f"An error occurred: {str(e)}"

def main():
    st.set_page_config(page_title="Excel Processor", layout="wide")
    
    st.title("Excel Data Processor")
    st.write("Upload your Excel file to process it based on the defined logic.")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])
    
    if uploaded_file is not None:
        if st.button("Process File"):
            with st.spinner('Processing...'):
                processed_df, output_filename, error = process_excel(uploaded_file)
                
                if error:
                    st.error(error)
                else:
                    st.success("File processed successfully!")
                    
                    st.subheader("Preview Result")
                    st.dataframe(processed_df)
                    
                    # Convert to CSV for download
                    csv = processed_df.to_csv(index=False).encode('utf-8')
                    
                    st.download_button(
                        label="Download CSV",
                        data=csv,
                        file_name=output_filename,
                        mime='text/csv',
                    )

if __name__ == "__main__":
    main()
