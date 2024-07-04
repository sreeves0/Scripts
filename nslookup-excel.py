# pip install pandas openpyxl
import pandas as pd
import socket

# Read the Excel file
file_path = 'C:\\Users\\Steven\\Desktop\\Tests\\Book.xlsx'
df = pd.read_excel(file_path)

# Initialize lists to store ratings and nslookup messages
ratings = []
nslookup_messages = []

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    hostname = row['hostname']
    expected_ip = row['ipv4']

    # Skip rows with empty IP address
    if pd.isna(expected_ip):
        message = f"Skipping {hostname} due to empty IP address"
        print(message)
        ratings.append("No tick")
        nslookup_messages.append(message)
        continue

    try:
        # Perform nslookup on hostname
        actual_ip = socket.gethostbyname(hostname)
        
        # Perform nslookup on expected_ip
        actual_hostname = socket.gethostbyaddr(expected_ip)[0]
        
        # Check if the IP addresses match
        if actual_ip == expected_ip and actual_hostname == hostname:
            message = f"Match: {hostname} - {expected_ip}"
            print(message)
            ratings.append("Tick")
        else:
            message = f"Mismatch: {hostname} - Expected: {expected_ip}, Got: {actual_ip}, Resolved Hostname: {actual_hostname}"
            print(message)
            ratings.append("No tick")
    except socket.gaierror as e:
        message = f"Error: Could not resolve {hostname} or {expected_ip} - {e}"
        print(message)
        ratings.append("No tick")
    except socket.herror as e:
        message = f"Error: Could not resolve {expected_ip} to a hostname - {e}"
        print(message)
        ratings.append("No tick")
    
    nslookup_messages.append(message)

# Add the ratings and nslookup messages to the DataFrame
df['rating'] = ratings
df['nslookup answer'] = nslookup_messages

# Save the DataFrame to a new Excel file
output_file_path = 'C:\\Users\\Steven\\Desktop\\Tests\\Book_with_ratings.xlsx'
df.to_excel(output_file_path, index=False)
