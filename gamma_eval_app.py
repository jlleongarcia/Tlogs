# -*- coding: utf-8 -*-

import streamlit as st
from pylinac import TrajectoryLog, load_log
import numpy as np
import os
import tempfile
from pathlib import Path
import pandas as pd
from datetime import datetime # To timestamp entries

# --- Configuration ---
RESULTS_FILE = r"\\srvhjerez\RADIOFISICA\QAs pacientes\Registro Tlogs.xlsx" # Define the results Excel file path

BODY_LOCATIONS = [ # Define the list of body locations
    "Cabeza y cuello",
    "Cerebral",
    "Canal anal",
    "Esófago",
    "SBRT Hepática",
    "Ginecológico",
    "Linfoma",
    "Mama libre con supra",
    "Mama libre sin supra",
    "Mama DIBH con supra",
    "Mama DIBH sin supra",
    "Piel",
    "Paliativo",
    "Próstata con cadenas",
    "Próstata sin cadenas",
    "Pulmón",
    "Recto",
    "Vejiga",
    "Radiocirugía",
    "Holocráneo",
    "SBRT Renal",
    "Extremidades",
    "SBRT Ósea",
    "SBRT Próstata",
    "SBRT Ganglionar",
    "SBRT Pulmón",
    "SBRT Páncreas",
    "Otros",
]

BODY_LOCATIONS.sort() # Sort the list alphabetically

# --- End Configuration ---


# --- Function to append data to Excel file with column width adjustment ---
def append_to_excel(filename, sheet_name, data_dict, column_widths=None):
    """Appends a dictionary as a new row to a specific sheet in an Excel file (.xlsx)
       and adjusts column widths.

    Args:
        filename (str): The path to the Excel file (.xlsx).
        sheet_name (str): The name of the sheet to append to.
        data_dict (dict): The dictionary representing the row to add.
        column_widths (dict, optional): Dictionary mapping column names to desired widths.
                                        If None, attempts basic auto-sizing. Defaults to None.
    """
    new_data_df = pd.DataFrame([data_dict])

    # Define standard column order
    standard_columns = [
        'Timestamp', 'Patient ID', 'Location', 'Threshold (%)', 'Gamma (%)',
        'Normalization', 'Gamma Octavius 4D (%)'
    ]
    # Reindex new data to ensure consistent column order
    new_data_df = new_data_df.reindex(columns=standard_columns)

    engine = 'xlsxwriter' # Use xlsxwriter engine

    try:
        all_sheets = {}
        file_exists = os.path.exists(filename)

        if file_exists:
            try:
                # Read all existing sheets
                with pd.ExcelFile(filename, engine='openpyxl') as xls:
                    all_sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

                if sheet_name in all_sheets:
                    existing_df = all_sheets[sheet_name]
                    # Ensure columns match before concatenating
                    existing_df = existing_df.reindex(columns=standard_columns)
                    updated_df = pd.concat([existing_df, new_data_df], ignore_index=True)
                else:
                    updated_df = new_data_df # Sheet doesn't exist
            except FileNotFoundError: # Handle case where file exists check passed but read fails
                 st.warning(f"File '{filename}' not found during read attempt. Creating a new file.")
                 updated_df = new_data_df
            except Exception as read_error:
                st.warning(f"Could not read existing file '{filename}'. It might be corrupted or incompatible. Creating a new file/sheet. Error: {read_error}")
                updated_df = new_data_df # Start fresh if reading fails
        else:
            # File doesn't exist
            updated_df = new_data_df

        # Update the specific sheet data in our dictionary
        all_sheets[sheet_name] = updated_df

        # Write all sheets back to the file using xlsxwriter
        with pd.ExcelWriter(filename, engine=engine) as writer:
            for sheet, df in all_sheets.items():
                # Ensure consistent column order when writing
                df_to_write = df.reindex(columns=standard_columns)
                df_to_write.to_excel(writer, sheet_name=sheet, index=False)

                # --- Column Width Adjustment ---
                workbook = writer.book
                worksheet = writer.sheets[sheet]

                if isinstance(column_widths, dict):
                    # Apply predefined widths
                    for idx, col_name in enumerate(standard_columns):
                        width = column_widths.get(col_name)
                        if width is not None:
                            worksheet.set_column(idx, idx, width)
                else:
                    # Basic auto-width calculation (adjust as needed)
                    for idx, col in enumerate(df_to_write):
                        series = df_to_write[col]
                        # Consider header length and max data length
                        max_len = max((
                            series.astype(str).map(len).max(), # Len of largest item
                            len(str(series.name))             # Len of column name/header
                            )) + 1 # Add a little buffer
                        # Clamp max width if needed, e.g., max_len = min(max_len, 50)
                        worksheet.set_column(idx, idx, max_len)
                # --- End Column Width Adjustment ---

        st.success(f"Data successfully saved to sheet '{sheet_name}' in {filename}")

    except ImportError as import_err:
         # Catch missing openpyxl or xlsxwriter
         st.error(f"Error: Required library not found. Please install it. Details: {import_err}")
    except PermissionError:
        st.error(f"Error: Permission denied when trying to write to {filename}. Check file/folder permissions and ensure the file is not open elsewhere.")
    except Exception as e:
        st.error(f"Error saving data to Excel file: {e}")

# --- End of function ---


# Define desired column widths (adjust these values as needed)
col_widths = {
    'Timestamp': 19, # Width for Timestamp column
    'Patient ID': 15,
    'Location': 25, # Wider for longer location names
    'Threshold (%)': 12,
    'Gamma (%)': 10,
    'Normalization': 12,
    'Gamma Octavius 4D (%)': 22
}


# --- Streamlit App UI ---

# Title of the Streamlit app
st.title('Gamma Evaluation')

# Radio button for single selection
option = st.radio("Choose an option:", ("VMAT", "SBRT", "SRS"), key="treatment_option")

# Determine parameters based on option
if option == "VMAT":
    dta = 0
    dd = 0.44
    res = 0.1
    sheet_name = "VMAT" # Target sheet for VMAT
else: # SBRT or SRS
    dta = 0
    dd = 0.69
    res = 0.1
    sheet_name = "SBRT-SRS" # Target sheet for SBRT/SRS

norm_option = st.radio("Choose local or global normalization:", ("Local", "Global"), key="norm_option")

if norm_option == "Local":
    normalize = False
else:
    normalize = True

thd_percent = st.number_input("Threshold (%): ", min_value=0.0, max_value=100.0, value=10.0, step=1.0, key="threshold_percent")
thd = thd_percent / 100.0 # Convert percentage to fraction for pylinac

# Display the gamma evaluation parameters
st.write('---')
st.subheader('Gamma Evaluation Parameters:')
st.write(f'*   Distance to Agreement (mm): {dta:.1f}') # Display DTA
st.write(f'*   Dose Difference (%): {dd:.2f}') # Display Dose Difference
st.write(f'*   Threshold (%): {thd * 100:.1f}') # Display Threshold
st.write(f'*   Resolution (mm): {res:.1f}') # Display Resolution
st.write(f'*   Normalization: {norm_option}') # Display Normalization
st.write('---')

# Load the trajectory log file
uploaded_files = st.file_uploader("Choose the .bin files for one patient", type=["bin"], accept_multiple_files=True, key="file_uploader")

# Initialize variables to store results outside the temp dir context
patient_id = None
all_gamma = None
selected_location = None
file_name_for_id = None # Store the name of the file used for ID

if uploaded_files:
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        
        # Dictionary to store file paths with filenames as keys
        file_paths = {}
        
        for uploaded_file in uploaded_files:
            # Save each uploaded file in the temp directory
            temp_file_path = os.path.join(temp_dir, uploaded_file.name)
            
            try:
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer()) # Use getbuffer() for uploaded files
                # Store the path in the dictionary
                file_paths[uploaded_file.name] = temp_file_path
                if file_name_for_id is None: # Get the name of the first uploaded file for ID extraction
                    file_name_for_id = uploaded_file.name
            except Exception as e:
                st.error(f"Error saving uploaded file {uploaded_file.name}: {e}")
                st.stop() # Stop execution if file saving fails
            
        # Check if any files were successfully saved
        if not file_paths:
            st.warning("No files were successfully processed.")
            st.stop()
        
        # --- Perform Analysis ---
        try:
            # Extract Patient ID from the first 12 characters of the first filename
            if file_name_for_id and len(file_name_for_id) >= 12:
                 patient_id = file_name_for_id[:12]
                 st.write(f"**Patient ID:** {patient_id}")
            else:
                 st.warning("Could not extract Patient ID (filename too short or no files).")
                 patient_id = "Unknown" # Assign a default value
        
            # Access files by their names using the dictionary
            file_name = st.selectbox("Select a file to process:", list(file_paths.keys()))
            selected_file_path = file_paths[file_name]

            # Load the trajectory log using the saved temporary file
            log = TrajectoryLog(selected_file_path)

            log.fluence.gamma.calc_map(distTA=dta, doseTA=dd, threshold=thd, resolution=res, normalize=normalize)
            st.write(f'Arc Gamma Passing Percentage: {log.fluence.gamma.pass_prcnt:.2f}')  # The gamma passing percentage of a single arc

            # Once Tlogs are individually analyzed, convert dictionary values to os.PathLike (Path objects)
            file_paths = {name: Path(path) for name, path in file_paths.items()}

            # Get all logs gamma
            all_logs = load_log(temp_dir)
            all_gamma = all_logs.avg_gamma_pct(distTA=dta, doseTA=dd, threshold=thd, resolution=res, normalize=normalize)
            st.write(f'Patient Gamma Passing Percentage: {all_gamma:.2f}')  # Gamma passing percentage for all arcs
        
        except Exception as analysis_error:
            st.error(f"An error occurred during log analysis: {analysis_error}")
            all_gamma = None # Ensure gamma is None if analysis fails
    
    # --- UI Elements for Saving (Outside temp dir context) ---
    if all_gamma is not None: # Only show saving options if analysis was successful
        st.write("---")
        st.subheader("Save Results")
        selected_location = st.selectbox("Select Body Location:", BODY_LOCATIONS, key="location_select")
        octavius_gamma = st.number_input("Octavius 4D Gamma (%): ", min_value=0.0, max_value=100.0, value=0.0, step=0.1, key="oct_gamma")

        # Submit button
        if st.button("Submit Results to Excel File", key="submit_button"):
            if patient_id and selected_location:
                # Prepare data dictionary for saving
                data_to_save = {
                    'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"), # Format timestamp
                    'Patient ID': patient_id,
                    'Location': selected_location,
                    'Threshold (%)': thd * 100,
                    'Gamma (%)': round(all_gamma, 2),
                    'Normalization': norm_option,
                    'Gamma Octavius 4D (%)': octavius_gamma
                }

                # Append data to the Excel file
                append_to_excel(RESULTS_FILE, sheet_name, data_to_save, column_widths=col_widths)
            else:
                st.warning("Cannot submit. Patient ID or Location is missing.")
else:
    st.info("Upload trajectory log (.bin) files to begin analysis.")
