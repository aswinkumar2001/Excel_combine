import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Streamlit app title
st.title("Combine Excel Files with Column Filtering")

# File uploaders for the two Excel files
database_file = st.file_uploader("Upload database.xlsx", type=["xlsx"])
discovered_file = st.file_uploader("Upload discovered.xlsx", type=["xlsx"])

if database_file is not None and discovered_file is not None:
    try:
        # Read Excel files
        database_sheets = pd.read_excel(database_file, sheet_name=None, engine='openpyxl')
        discovered_sheets = pd.read_excel(discovered_file, sheet_name=None, engine='openpyxl')

        # Create an in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Process database.xlsx sheets
            for sheet_name, df in database_sheets.items():
                new_sheet_name = f"{sheet_name}_database"[:31]  # Excel sheet names max 31 chars
                try:
                    # Filter for required columns
                    required_columns = ['Path', 'Name', 'Object ID']
                    available_columns = [col for col in required_columns if col in df.columns]
                    if not available_columns:
                        st.warning(f"No required columns (Path, Name, Object ID) found in {sheet_name} from database.xlsx. Skipping.")
                        continue
                    df_filtered = df[available_columns]
                    df_filtered.to_excel(writer, sheet_name=new_sheet_name, index=False)
                    st.write(f"Processed: {sheet_name} from database.xlsx as {new_sheet_name} with columns {available_columns}")
                except Exception as e:
                    st.warning(f"Error processing {sheet_name} from database.xlsx: {str(e)}")

            # Process discovered.xlsx sheets
            for sheet_name, df in discovered_sheets.items():
                new_sheet_name = f"{sheet_name}_discovered"[:31]  # Excel sheet names max 31 chars
                try:
                    # Filter for required columns
                    required_columns = ['Object Name', 'Object ID']
                    available_columns = [col for col in required_columns if col in df.columns]
                    if not available_columns:
                        st.warning(f"No required columns (Object Name, Object ID) found in {sheet_name} from discovered.xlsx. Skipping.")
                        continue
                    df_filtered = df[available_columns]
                    df_filtered.to_excel(writer, sheet_name=new_sheet_name, index=False)
                    st.write(f"Processed: {sheet_name} from discovered.xlsx as {new_sheet_name} with columns {available_columns}")
                except Exception as e:
                    st.warning(f"Error processing {sheet_name} from discovered.xlsx: {str(e)}")

        # Check if any sheets were written
        if writer.sheets:
            # Prepare file for download
            output.seek(0)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"combined_excel_{timestamp}.xlsx"

            st.success(f"Combined {len(writer.sheets)} sheets into one Excel file.")
            st.download_button(
                label="Download Combined Excel File",
                data=output,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No sheets were processed. Check if the Excel files contain the required columns.")

    except Exception as e:
        st.error(f"Error processing Excel files: {str(e)}")
else:
    st.info("Please upload both database.xlsx and discovered.xlsx files to proceed.")
