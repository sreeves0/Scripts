import pandas as pd
import socket

# Read input Excel file into pandas DataFrame
input_file = 'C:\\Users\\WE21927\\DJProject\\Project02\\checkpoint-objects1.xlsx'  # Replace with your input Excel file name
output_file = 'C:\\Users\\WE21927\\DJProject\\Project02\\output.xlsx'  # Replace with your output Excel file name
df = pd.read_excel(input_file)
# Create empty list to store results
results = []

# Function to perform DNS lookup and print messages
def perform_dns_lookup_and_print(hostname):
    try:
        ip_address = socket.gethostbyname(hostname)
        print(f"Domain exists: {hostname} ({ip_address})")
        return 10  # Hostname resolved successfully
    except socket.gaierror:
        print(f"Domain does not exist: {hostname}")
        return 0   # Hostname not resolved (NXDOMAIN)
# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    if pd.notna(row['IPv4 address']):  # Check if IPv4 Address is not empty
        name_result = perform_dns_lookup_and_print(row['Name'])
        ip_result = perform_dns_lookup_and_print(row['IPv4 address'])
        
        # Determine the score based on the results
        if name_result == 10 or ip_result == 10:
            results.append(10)
        else:
            results.append(0)
    else:
        print(f"Skipping row {index + 1}: IPv4 Address is empty")
        results.append(None)  # If IPv4 Address is empty, append None
# Add results list as a new column 'Score' in the DataFrame
df['Score'] = results
# Save the updated DataFrame to a new Excel file
df.to_excel(output_file, index=False)
print(f"Output saved to {output_file}")
