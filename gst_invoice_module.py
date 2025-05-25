import os
import pandas as pd
from datetime import datetime, timedelta
from generateInvoice import generate_invoice
import numpy as np

# Create a GST invoice number based on the logic shared and paste it in the GST template as per the below mapping
def create_GSTInvoiceno(row, df1, sequence_number):
    column_name = 'AGENT_STATE' # The column containing the agent's state
    
    # Get the state name from the first row of the specified column
    first_row_state = row[column_name]
    df1.columns = df1.columns.str.strip()

    # Filter rows in the second DataFrame where the state matches the state from the first row
    # and select 'gstno' and 'code' columns
    filtered_df = df1.loc[df1['State/UT'] == first_row_state, ['GSTIN No', 'Statecode']]
    
    if not filtered_df.empty:
        # Get the state code from the filtered DataFrame
        state_code = filtered_df.iloc[0]['Statecode']
        gst_no = filtered_df.iloc[0]['GSTIN No']
        get_ISSdate = row['ISSDATE']

        # Define the constant string based on whether GST number exists
        constant_string = "IGB" if pd.notnull(gst_no) else "IGC"
        
        # Get the current month and year in MMYY format
        current_date = datetime.now().strftime("%m%y")
        
        # Return the GST invoice number
        gst_invoice_no = f"{state_code}/{constant_string}"
        return gst_invoice_no, gst_no
    
    else:
        print("No matching state found in the GST master file.")
        return None, None

# Function to convert Excel serial numbers date to the formatted date as MMYY  
def excel_serial_to_date(serial_number):
    base_date = datetime(1899, 12, 30)  # Excel's base date
    
    # Convert numpy int64 to standard Python int if needed
    if isinstance(serial_number, np.int64):
        serial_number = int(serial_number)
    
    actual_date = base_date + timedelta(days=serial_number)
    return actual_date.strftime("%m%y") # Return the formatted date as MMYY

# Function to check if a file exists and process it
def check_file_exists(file_path):
    try:
        if os.path.isfile(file_path):
            print(f"The file exists in the path.")
            sequence_number = 1 #Set the sequence number used for naming file
            
            # Read the CSV file using pandas and dataframe(df) has all the data from inputfile
            df = pd.read_csv(file_path)
            df.columns = df.columns.str.strip()
            
            # Filter the DataFrame for transactions from previous day from current day
            previous_date = datetime.now() - timedelta(days=1)
            df_input = df[df['TRANSACTION_DATE'] == previous_date.strftime('%d-%m-%Y')]
            
            
            for index, row in df_input.iterrows():
                df1 = pd.read_excel("C:/Users/Admin/OneDrive/Desktop/HDFL_Invoice/HDFC_Life/HDFC/Invoice_Format_HDFC_GST_number_master.xlsx", sheet_name="state code")
                gst_invoice_no, gst_no = create_GSTInvoiceno(row, df1, sequence_number)
                
                if gst_invoice_no:
                    # Extract relevant details for invoice creation
                    invoice_details = row[['CONTRACT_NO', 'OWNER_CLIENT_NO', 'CLIENT_FULLNAME', 'INSURED_FROM(RCD date)', 'ZGSTNO', 
                                           'CLTADDR01', 'CLTADDR02', 'CLTADDR03', 'CLTADDR04', 'CLTADDR05', 'CLTPCODE', 'NETPREMIUM', 
                                           'ISSDATE', 'ZHSGSTAMT', 'ZHCGSTAMT', 'ZHIGSTAMT']]
                    
                    invoice_details['GSTIN no'] = gst_no
                    invoice_details['GST_INVOICE_NO'] = gst_invoice_no
                    invoice_details['TODAYS_DATE'] = datetime.now().strftime('%d-%m-%Y')
                    Risk_commence_date=str(invoice_details['INSURED_FROM(RCD date)'])

                    # Update the risk commence date format
                    Risk_commence_date=Risk_commence_date.replace('-','/')
                    invoice_details['INSURED_FROM(RCD date)']=Risk_commence_date

                    # Assign the sequence number to the invoice
                    invoice_details['SequenceNo']=sequence_number
                    # Generate the invoice
                    generate_invoice(invoice_details)

                    # Increment the sequence number for the next invoice
                    sequence_number += 1 
                    if sequence_number >= 9999: # Check if the sequence number exceeds the limit
                     raise ValueError("Sequence has exceeded 9999. Please submit a change request (CR).")
                    
            print(f"Generated GST Invoice")
        else:
            print(f"The file does not exist in the path.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Define the path to the input file
path = "C:/Users/Admin/OneDrive/Desktop/HDFL_Invoice/HDFC_Life/HDFC/"

# Input file name : SR_238(default)_MMM_Date(YYYYMMDD) (current date T day)
current_date = datetime.now().date()
formatted_date = current_date.strftime("_%b_%Y%m%d")
file_path = path + "SR_238" + formatted_date + ".csv"

check_file_exists(file_path)
