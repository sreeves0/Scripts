import pandas as pd
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

# Read the Excel file
file_path = 'C:\\CyberProjects\\Scripts\\Book.xlsx'
df = pd.read_excel(file_path)
# Rename the columns
df.columns = ['Name', 'IPv4 address'] + list(df.columns[2:])
# Initialize lists to store ratings
ratings_col1 = [0] * len(df)
ratings_col2 = [0] * len(df)
def nslookup(domain, index, col):
    try:
        result = subprocess.run(['nslookup', domain], capture_output=True, text=True)
        print(f"nslookup for {domain}:\n{result.stdout}")
        if "Name:" in result.stdout:
            return (index, col, 5)
        else:
            return (index, col, 0)
    except Exception as e:
        print(f"Error performing nslookup on {domain}: {e}")
        return (index, col, 0)
# Perform nslookup on the first and second columns in parallel
with ThreadPoolExecutor(max_workers=10) as executor:
    futures = []
    for i, (domain1, domain2) in enumerate(zip(df['Name'], df['IPv4 address'])):
        futures.append(executor.submit(nslookup, domain1, i, 'col1'))
        if pd.notna(domain2):
            futures.append(executor.submit(nslookup, domain2, i, 'col2'))
    for future in as_completed(futures):
        index, col, rating = future.result()
        if col == 'col1':
            ratings_col1[index] = rating
        else:
            ratings_col2[index] = rating
# Add ratings to the DataFrame
df['Rating_Col1'] = ratings_col1
df['Rating_Col2'] = ratings_col2
# Create a new column that adds the values from Rating_Col1 and Rating_Col2
df['Total_Rating'] = df['Rating_Col1'] + df['Rating_Col2']
# Save the updated DataFrame to a new Excel file
df.to_excel('C:\\CyberProjects\\Scripts\\Book-result.xlsx', index=False)
