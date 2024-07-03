# pip install pandas openpyxl
import pandas as pd
import socket

# Read the Excel file
file_path = 'C:\Users\Steven\Desktop\Tests\Book.xlsx'
df = pd.read_excel(file_path)

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    hostname = row['hostname']
    expected_ip = row['ipv4']

    # Skip rows with empty IP address
    if pd.isna(expected_ip):
        print(f"Skipping {hostname} due to empty IP address")
        continue

    try:
        # Perform nslookup on hostname
        actual_ip = socket.gethostbyname(hostname)
        
        # Perform nslookup on expected_ip
        actual_hostname = socket.gethostbyaddr(expected_ip)[0]
        
        # Check if the IP addresses match
        if actual_ip == expected_ip and actual_hostname == hostname:
            print(f"Match: {hostname} - {expected_ip}")
        else:
            print(f"Mismatch: {hostname} - Expected: {expected_ip}, Got: {actual_ip}, Resolved Hostname: {actual_hostname}")
    except socket.gaierror as e:
        print(f"Error: Could not resolve {hostname} or {expected_ip} - {e}")
    except socket.herror as e:
        print(f"Error: Could not resolve {expected_ip} to a hostname - {e}")
