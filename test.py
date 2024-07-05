import pandas as pd
import subprocess

def nslookup(value):
    try:
        result = subprocess.run(['nslookup', value], capture_output=True, text=True)
        if 'NXDOMAIN' in result.stderr or 'non-existent' in result.stdout:
            return 0
        else:
            return 10
    except Exception as e:
        print(f"Error performing nslookup for {value}: {e}")
        return 0

def main(excel_file):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Check if required columns exist
    if 'Names' not in df.columns or 'IPv4 Address' not in df.columns:
        print("The Excel file must contain 'Names' and 'IPv4 Address' columns.")
        return

    # Perform nslookup and assign scores
    df['Names Score'] = df['Names'].apply(nslookup)
    
    def ipv4_nslookup(ip):
        if pd.isna(ip) or ip.strip() == '':
            return None
        return nslookup(ip)

    df['IPv4 Address Score'] = df['IPv4 Address'].apply(ipv4_nslookup)

    # Save the results to a new Excel file
    output_file = 'nslookup_results.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

if __name__ == "__main__":
    # Replace 'your_excel_file.xlsx' with your actual Excel file name
    main('your_excel_file.xlsx')