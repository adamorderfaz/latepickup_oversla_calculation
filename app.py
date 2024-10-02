import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import timedelta

# Function to process the uploaded file
def process_data(df):
    # Select specific columns
    selected_columns = ['No', 'Order ID', 'Shipment ID', 'Tracking ID', 'Courier',
                        'Courier Service', 'Shipment Type', 'Tracking Status',
                        'Created at', 'Dispatch at', 'Pickup Date',
                        'min SLA', 'max SLA', 'Delivered at 2']

    # Create a new dataframe with only the selected columns
    selected_data = df[selected_columns]

    # Convert the 'Created at', 'Dispatched at', and 'Pickup Date' columns to datetime format
    selected_data['Created at'] = pd.to_datetime(selected_data['Created at'], format='%d %b %y %H:%M', errors='coerce')
    selected_data['Dispatch at'] = pd.to_datetime(selected_data['Dispatch at'], format='%d %b %y %H:%M',
                                                  errors='coerce')
    selected_data['Pickup Date'] = pd.to_datetime(selected_data['Pickup Date'], format='%d %b %y %H:%M',
                                                  errors='coerce')
    selected_data['Delivered at 2'] = pd.to_datetime(selected_data['Delivered at 2'], format='%d %b %y',
                                                     errors='coerce')

    # Define the cutoff time for the day
    cutoff_time = pd.to_datetime("15:00:00").time()

    # Function to determine if the pickup was late or not
    def determine_late_pickup(row):
        if pd.isnull(row['Dispatch at']) or pd.isnull(row['Pickup Date']):
            return 'Non-Late Pickup'

        request_time = row['Pickup Date'].time()  # The time user requested pickup
        actual_pickup_datetime = row['Dispatch at']  # The datetime when courier picked up the package
        request_datetime = row['Pickup Date']  # The datetime when user requested the pickup

        # Calculate the time difference between pickup request and actual dispatch
        time_difference = actual_pickup_datetime - request_datetime

        # If the time difference exceeds 24 hours, mark as Late Pickup
        if time_difference > pd.Timedelta(days=1):
            return 'Late Pickup'

        # Apply the original late pickup logic
        if request_time >= cutoff_time and actual_pickup_datetime.date() > request_datetime.date():
            return 'Non-Late Pickup'

        if request_time < cutoff_time and actual_pickup_datetime.date() == request_datetime.date():
            return 'Non-Late Pickup'

        if request_time < cutoff_time and actual_pickup_datetime.date() > request_datetime.date():
            return 'Late Pickup'

        return 'Non-Late Pickup'

    # Apply the function to determine late or non-late pickup
    selected_data['Late Pickup'] = selected_data.apply(determine_late_pickup, axis=1)

    # Calculate 'Tanggal Max' as 'Dispatch at' + 'max SLA' (in days)
    selected_data['Tanggal Max'] = selected_data['Dispatch at'] + pd.to_timedelta(selected_data['max SLA'], unit='days')

    # Calculate 'Over SLA' as the difference between 'Tanggal Max' and 'Delivered at 2' (in days)
    selected_data['Over SLA'] = (selected_data['Tanggal Max'] - selected_data['Delivered at 2']).dt.days
    selected_data['Over SLA'] = selected_data['Over SLA'].astype(int, errors='ignore')

    # Get the current date as 'Today'
    today_date = pd.to_datetime('today').normalize()

    # Add the 'Today' column to the dataframe
    selected_data['Today'] = today_date

    # Calculate 'Over SLA IP'
    selected_data['Over SLA IP'] = (selected_data['Today'] - selected_data['Tanggal Max']).dt.days

    return selected_data

# Function to convert dataframe to Excel
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Processed Data')
    writer.close()  # Correct method to close the writer
    processed_data = output.getvalue()
    return processed_data

# Streamlit UI
st.title("Shipment Data Processor")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

# Initialize session state for preventing refresh
if "df" not in st.session_state:
    st.session_state.df = None

if uploaded_file:
    # Read the uploaded Excel file
    df = pd.read_excel(uploaded_file)

    # Process the data
    processed_df = process_data(df)

    # Store processed dataframe in session state
    st.session_state.df = processed_df

    # Display the processed dataframe
    st.write("Processed Data:")
    st.dataframe(processed_df)

    # Download button for processed data
    processed_excel = to_excel(st.session_state.df)
    st.download_button(
        label="Download Processed Data",
        data=processed_excel,
        file_name="processed_shipment_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.write("Please upload an Excel file to process.")
