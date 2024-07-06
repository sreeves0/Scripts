import pandas as pd
import subprocess

# Read the Excel file
file_path = 'C:\\CyberProjects\\Scripts\\Book.xlsx'
df = pd.read_excel(file_path)
# Rename the columns
df.columns = ['Name', 'IPv4 address'] + list(df.columns[2:])
# Initialize lists to store ratings
ratings_col1 = []
ratings_col2 = []
# Perform nslookup on the first and second columns
for domain1, domain2 in zip(df['Name'], df['IPv4 address']):
    try:
        result1 = subprocess.run(['nslookup', domain1], capture_output=True, text=True)
        print(f"nslookup for {domain1}:\n{result1.stdout}")
        if "Name:" in result1.stdout:
            ratings_col1.append(5)
        else:
            ratings_col1.append(0)
    except Exception as e:
        print(f"Error performing nslookup on {domain1}: {e}")
        ratings_col1.append(0)
    try:
        result2 = subprocess.run(['nslookup', domain2], capture_output=True, text=True)
        print(f"nslookup for {domain2}:\n{result2.stdout}")
        if "Name:" in result2.stdout:
            ratings_col2.append(5)
        else:
            ratings_col2.append(0)
    except Exception as e:
        print(f"Error performing nslookup on {domain2}: {e}")
        ratings_col2.append(0)
# Add ratings to the DataFrame
df['Rating_Col1'] = ratings_col1
df['Rating_Col2'] = ratings_col2
# Create a new column that adds the values from Rating_Col1 and Rating_Col2
df['Total_Rating'] = df['Rating_Col1'] + df['Rating_Col2']
# Save the updated DataFrame to a new Excel file
df.to_excel('C:\\CyberProjects\\Scripts\\Book-result.xlsx', index=False)
