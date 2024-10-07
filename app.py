import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
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

    # Define a helper function for flexible date parsing
    def parse_dates(date_str):
        formats = ['%d/%m/%Y %H:%M:%S', '%d-%m-%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d-%m-%Y %H:%M', '%Y-%m-%d %H:%M:%S',
                   '%Y-%m-%d %H:%M']
        for fmt in formats:
            try:
                return pd.to_datetime(date_str, format=fmt, dayfirst=True)
            except (ValueError, TypeError):
                continue
        return pd.NaT  # Return NaT if no format matches

    # Try to convert the columns to datetime and handle errors
    try:
        # Parse dates using the helper function for flexible date formats
        selected_data['Created at'] = selected_data['Created at'].apply(parse_dates)
        selected_data['Dispatch at'] = selected_data['Dispatch at'].apply(parse_dates)
        selected_data['Pickup Date'] = selected_data['Pickup Date'].apply(parse_dates)
        # Ensure 'Delivered at' is treated as datetime, including null values
        selected_data['Delivered at'] = selected_data['Delivered at'].apply(parse_dates)
    except Exception as e:
        st.error(f"Error in datetime conversion: {str(e)}")
        return None

    # Report missing or invalid datetime values, but do not remove them
    if selected_data[['Created at', 'Dispatch at', 'Pickup Date', 'Delivered at']].isnull().any().any():
        st.warning(
            "Some rows in Dispatch at / Delivered at are empty, so they are marked as 'Dispatch at Kosong' or 'Delivered at Kosong'.")

    # Apply the function to determine late or non-late pickup and 'Days Late'
    selected_data[['Late Pickup', 'Days Late']] = selected_data.apply(lambda row: determine_late_pickup(row), axis=1,
                                                                      result_type='expand')

    # Calculate 'Tanggal Max' as 'Dispatch at' + 'max SLA' (in days), or mark 'Dispatch at Kosong'
    selected_data['Tanggal Max'] = selected_data.apply(
        lambda row: row['Dispatch at'] + pd.to_timedelta(row['max SLA'], unit='days')
        if pd.notnull(row['Dispatch at'])
        else pd.NaT, axis=1  # Use NaT instead of 'Dispatch at Kosong'
    )

    # Calculate 'Over SLA' as the difference between 'Tanggal Max' and 'Delivered at', or mark 'Dispatch at Kosong' or 'Delivered at Kosong'
    selected_data['Over SLA'] = selected_data.apply(
        lambda row: (row['Tanggal Max'] - row['Delivered at']).days
        if pd.notnull(row['Tanggal Max']) and pd.notnull(row['Delivered at'])
        else 'Delivered at Kosong' if pd.isnull(row['Delivered at']) else 'Dispatch at Kosong', axis=1
    )

    # Get the current date as 'Today' and format it as 'DD-MM-YYYY'
    today_date = datetime.today().strftime('%d-%m-%Y')

    # Set Today to match the desired format and calculate Over SLA IP
    selected_data['Today'] = pd.to_datetime(today_date, format='%d-%m-%Y')

    # Calculate 'Over SLA IP' strictly based on the date (not time)
    selected_data['Over SLA IP'] = selected_data.apply(
        lambda row: (selected_data['Today'].iloc[0] - row['Tanggal Max'].normalize()).days
        if pd.notnull(row['Tanggal Max']) and row['Tanggal Max'] != pd.NaT
        else 'Dispatch at Kosong', axis=1
    )

    # Format the datetime columns to the desired format 'DD-MM-YYYY'
    selected_data['Today'] = selected_data['Today'].dt.strftime('%d-%m-%Y')
    datetime_columns = ['Created at', 'Dispatch at', 'Pickup Date', 'Delivered at', 'Tanggal Max']
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

# File uploader, now supports both CSV and XLSX
uploaded_file = st.file_uploader("Upload your file (XLSX or CSV)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # Check the file type and read accordingly
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)

    except Exception as e:
        st.error(f"Failed to read the uploaded file: {str(e)}")
        df = None

    if df is not None:
        # Process the data
        processed_df = process_data(df)
        processed_df = processed_df.drop(columns = 'Unnamed: 0')

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
    st.write("Please upload an Excel or CSV file to process.")
