import streamlit as st
import pandas as pd
import io

# --- Main App Logic ---

def process_files(uploaded_files):
    """
    Reads uploaded Excel files, processes them according to the original logic,
    and returns the final DataFrame.
    """
    dfs = []
    for file_obj in uploaded_files:
        try:
            source_workbook = pd.ExcelFile(file_obj)
            for i, sheet_name in enumerate(source_workbook.sheet_names):
                df = source_workbook.parse(sheet_name, header=None)
                
                # Remove the first two rows from the first sheet only
                if i == 0:
                    df = df.iloc[2:]
                
                dfs.append(df)
        except Exception as e:
            st.warning(f"Could not process file: {file_obj.name}. Error: {e}")
            continue # Move to the next file if one fails

    if not dfs:
        return None

    # Concatenate all DataFrames
    stacked_df = pd.concat(dfs, ignore_index=True)

    # --- Data Cleaning and Transformation (from original script) ---
    
    # 1. Drop rows with NaN in the 13th column (index 12)
    stacked_df = stacked_df.dropna(subset=[stacked_df.columns[12]])

    # 2. Convert specific columns to numeric, coercing errors to NaN
    stacked_df[stacked_df.columns[12]] = pd.to_numeric(stacked_df[stacked_df.columns[12]], errors='coerce')
    stacked_df[stacked_df.columns[18]] = pd.to_numeric(stacked_df[stacked_df.columns[18]], errors='coerce')

    # 3. Define and set new column labels
    new_column_labels = [
        "ลำดับ", "วันที่จ่ายยา", "เวลา", "เลขที่เอกสาร", "VN / AN",
        "HN", "ชื่อ", "อายุ", "สิทธิ์", "แพทย์", "Clinic",
        "Ward", "Material", "รายการยา", "จำนวน", "หน่วย",
        "ราคาขายR", "ราคารวม", "Store"
    ]
    # Ensure the dataframe has the correct number of columns before renaming
    if len(stacked_df.columns) == len(new_column_labels):
        stacked_df.columns = new_column_labels
    else:
        # Handle cases where column counts don't match, maybe by trimming or erroring
        # For now, we'll just show an error and stop.
        st.error(f"Column count mismatch. Expected {len(new_column_labels)} but got {len(stacked_df.columns)}.")
        return None

    # 4. Define valid values for filtering
    valid_material_values = [
        1400000010, 1400000020, 1400000021, 1400000025, 1400000029, 
        1400000030, 1400000040, 1400000044, 1400000052, 1400000053, 
        1400000055, 1400000098, 1400000099, 1400000148, 1400000187, 
        1400000201, 1400000220, 1400000221, 1400000228, 1400000247, 
        1400000264, 1400000068, 1400000069, 1400000093, 1400000106, 
        1400000113, 1400000115, 1400000116, 1400000118, 1400000124, 
        1400000126, 1400000130, 1400000165, 1400000166, 1400000167, 
        1400000168, 1400000169, 1400000170, 1400000171, 1400000172, 
        1400000194, 1400000284, 1400000288, 1400000294, 1400000295, 
        1400000331, 1400000335, 1400000344, 1400000345, 1400000265
    ]

    # 5. Filter the DataFrame based on the "Material" column
    merged_df = stacked_df[stacked_df["Material"].isin(valid_material_values)].copy()

    # 6. Select and reorder the final columns
    final_cols = ['วันที่จ่ายยา', 'VN / AN', 'HN', 'ชื่อ', 'สิทธิ์', "แพทย์", 'Material', 'รายการยา', 'จำนวน']
    merged_df = merged_df[final_cols]

    return merged_df

# --- Streamlit UI ---

st.set_page_config(layout="wide")
st.title("📂 Excel File Merger and Processor")
st.write("This application merges all sheets from multiple `.xls` files, cleans the data, and filters it based on predefined criteria.")

# File uploader allows multiple files of type xls
uploaded_files = st.file_uploader(
    "Choose .xls files",
    type="xls",
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📁 You have uploaded {len(uploaded_files)} files.")

    # Process files when button is clicked
    if st.button("🚀 Process Files"):
        with st.spinner("Processing... Please wait."):
            final_df = process_files(uploaded_files)

        if final_df is not None:
            st.success("✅ Processing complete!")
            
            st.subheader("Filtered Data Preview")
            st.dataframe(final_df)
            
            # --- Download Button ---
            # Convert DataFrame to an in-memory Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Raw')
            
            # It's important to get the value of the BytesIO object
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Download Processed File (J2.xlsx)",
                data=excel_data,
                file_name="J2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Processing failed. No data was generated. Please check your files.")
