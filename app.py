import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import timedelta
from PIL import Image

# List of required columns in the desired order
required_columns = [
    'No', 'Order ID', 'Shipment ID', 'Tracking ID', 'Courier', 'Courier Service',
    'Shipment Type', 'Tracking Status', 'Created at', 'Dispatch at', 'Pickup Date',
    'min SLA', 'max SLA', 'Delivered at'
]


# Function to validate if all required columns are present
def validate_columns(df):
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return False, missing_columns
    return True, None


# Function to determine if the pickup was late or not, and how many days late
def determine_late_pickup(row):
    if pd.isnull(row['Dispatch at']):
        return 'Dispatch at Kosong', 'Dispatch at Kosong'

    if pd.isnull(row['Pickup Date']):
        return 'Non-Late Pickup', 0

    request_time = row['Pickup Date'].time()  # The time user requested pickup
    actual_pickup_datetime = row['Dispatch at']  # The datetime when courier picked up the package
    request_datetime = row['Pickup Date']  # The datetime when user requested the pickup

    # Calculate the time difference between pickup request and actual dispatch
    time_difference = actual_pickup_datetime - request_datetime

    # If the time difference exceeds 24 hours, mark as Late Pickup and return the number of days late
    if time_difference > pd.Timedelta(days=1):
        return 'Late Pickup', time_difference.days

    # Apply the original late pickup logic
    if request_time >= pd.to_datetime("15:00:00").time() and actual_pickup_datetime.date() > request_datetime.date():
        return 'Non-Late Pickup', 0

    if request_time < pd.to_datetime("15:00:00").time() and actual_pickup_datetime.date() == request_datetime.date():
        return 'Non-Late Pickup', 0

    if request_time < pd.to_datetime("15:00:00").time() and actual_pickup_datetime.date() > request_datetime.date():
        return 'Late Pickup', 1

    return 'Non-Late Pickup', 0


# Function to process the uploaded file
def process_data(df):
    # Validate columns
    is_valid, missing_columns = validate_columns(df)
    if not is_valid:
        st.error(f"The following columns are missing: {', '.join(missing_columns)}")
        return None

    # Select all columns (no subset), as requested by user
    selected_data = df.copy()

    # Try to convert the columns to datetime and handle errors
    try:
        selected_data['Created at'] = pd.to_datetime(selected_data['Created at'], format='%d %b %y %H:%M',
                                                     errors='coerce')
        selected_data['Dispatch at'] = pd.to_datetime(selected_data['Dispatch at'], format='%d %b %y %H:%M',
                                                      errors='coerce')
        selected_data['Pickup Date'] = pd.to_datetime(selected_data['Pickup Date'], format='%d %b %y %H:%M',
                                                      errors='coerce')
        selected_data['Delivered at'] = pd.to_datetime(selected_data['Delivered at'], format='%d %b %y',
                                                       errors='coerce')
    except Exception as e:
        st.error(f"Error in datetime conversion: {str(e)}")
        return None

    # Report missing or invalid datetime values, but do not remove them
    if selected_data[['Created at', 'Dispatch at', 'Pickup Date', 'Delivered at']].isnull().any().any():
        st.warning(
            "Beberapa baris data pada kolom Dispatch at / Delivered at kosong, jadi perhitungan diabaikan dan diberikan penanda 'Dispatch at Kosong' atau 'Delivered at Kosong'")

    # Apply the function to determine late or non-late pickup and 'Days Late'
    selected_data[['Late Pickup', 'Days Late']] = selected_data.apply(lambda row: determine_late_pickup(row), axis=1,
                                                                      result_type='expand')

    # Calculate 'Tanggal Max' as 'Dispatch at' + 'max SLA' (in days), or mark 'Dispatch at Kosong'
    selected_data['Tanggal Max'] = selected_data.apply(
        lambda row: row['Dispatch at'] + pd.to_timedelta(row['max SLA'], unit='days')
        if pd.notnull(row['Dispatch at'])
        else 'Dispatch at Kosong', axis=1
    )

    # Calculate 'Over SLA' as the difference between 'Tanggal Max' and 'Delivered at', or mark 'Dispatch at Kosong' or 'Delivered at Kosong'
    selected_data['Over SLA'] = selected_data.apply(
        lambda row: (row['Tanggal Max'] - row['Delivered at']).days
        if pd.notnull(row['Tanggal Max']) and pd.notnull(row['Delivered at']) and row[
            'Tanggal Max'] != 'Dispatch at Kosong'
        else 'Delivered at Kosong' if pd.isnull(row['Delivered at']) else 'Dispatch at Kosong', axis=1
    )

    # Get the current date as 'Today'
    today_date = pd.to_datetime('today').normalize()
    selected_data['Today'] = today_date

    # Calculate 'Over SLA IP' only for statuses that are not 'Delivered', 'Returned', or 'Waiting for Pickup'
    conditions = ~selected_data['Tracking Status'].isin(['Delivered', 'Returned', 'Waiting for Pickup'])

    # Calculate 'Over SLA IP' or mark 'Dispatch at Kosong' or 'Delivered at Kosong'
    selected_data['Over SLA IP'] = selected_data.apply(
        lambda row: (row['Today'] - row['Tanggal Max']).days
        if pd.notnull(row['Tanggal Max']) and row['Tanggal Max'] != 'Dispatch at Kosong' and conditions[row.name]
        else 'Delivered at Kosong' if pd.isnull(row['Delivered at']) else 'Dispatch at Kosong', axis=1
    )

    # Format the datetime columns to the desired format 'DD-MM-YYYY hh:mm:ss'
    datetime_columns = ['Created at', 'Dispatch at', 'Pickup Date', 'Delivered at', 'Tanggal Max', 'Today']
    for col in datetime_columns:
        selected_data[col] = pd.to_datetime(selected_data[col], errors='coerce').dt.strftime('%d-%m-%Y %H:%M:%S')

    # Rearrange columns based on the required order
    required_columns_order = [
        'No', 'Order ID', 'Shipment ID', 'Tracking ID', 'Courier', 'Courier Service',
        'Shipment Type', 'Tracking Status', 'Created at', 'Dispatch at', 'Pickup Date',
        'min SLA', 'max SLA', 'Delivered at', 'Late Pickup', 'Days Late',
        'Tanggal Max', 'Over SLA', 'Today', 'Over SLA IP'
    ]
    # Add remaining columns that were not specified in the required order
    remaining_columns = [col for col in selected_data.columns if col not in required_columns_order]
    cols_in_order = required_columns_order + remaining_columns
    selected_data = selected_data[cols_in_order]

    return selected_data


# Function to convert dataframe to Excel and treat specific columns as text (to avoid scientific notation)
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # Ensure columns like 'Tracking ID' are treated as text to avoid scientific notation
    df['Tracking ID'] = df['Tracking ID'].astype(str)  # Convert Tracking ID to string

    df.to_excel(writer, index=False, sheet_name='Processed Data')

    workbook = writer.book
    worksheet = writer.sheets['Processed Data']

    # Set format for 'Tracking ID' to ensure it remains text and doesn't appear in scientific notation
    text_format = workbook.add_format({'num_format': '@'})  # '@' means text format
    worksheet.set_column('D:D', None, text_format)  # Assuming 'Tracking ID' is in column D (adjust if necessary)

    writer.close()
    processed_data = output.getvalue()
    return processed_data


# Load images as icon
icon_image = Image.open("orderfaz.jpeg")

# Page Config
st.set_page_config(page_title="Orderfaz - Late Pickup & SLA Calculation", page_icon=icon_image, layout="wide")

# Streamlit UI
st.title("Orderfaz - Late Pickup & Over SLA Calculation")

# Note
st.success('CATATAN : Pastikan untuk input data sesuai dengan format yang sesuai, jangan ada perubahan')

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

if uploaded_file:
    # Read the uploaded Excel file
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Failed to read the uploaded file: {str(e)}")
        df = None

    if df is not None:
        # Process the data
        processed_df = process_data(df)

        if processed_df is not None:
            # Display the processed dataframe
            st.write("Processed Data:")
            st.dataframe(processed_df)

            # Download button for processed data
            processed_excel = to_excel(processed_df)
            st.download_button(
                label="Download Processed Data",
                data=processed_excel,
                file_name="processed_shipment_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.write("Please upload an Excel file to process.")
