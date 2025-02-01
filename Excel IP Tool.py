import pandas as pd
import subprocess
import os
import tempfile
import shutil
import time  

def get_excel_file():
    while True:
        file_path = input("Enter Excel File Path OR Name: ").strip()
        if not os.path.exists(file_path):
            print(f"❌ Typo or The File Doesnt Exist! ❌: {file_path}")
            continue
        try:
            open(file_path, 'a').close()
            return file_path
        except PermissionError:
            print(f"❌ You Left Excel Open! ❌: {file_path}")

def save_data(df, excel_file):
    temp_dir = tempfile.gettempdir()
    temp_file = os.path.join(temp_dir, os.path.basename(excel_file))
    
    try:
        df.to_excel(temp_file, index=False, engine='openpyxl')
        shutil.copy(temp_file, excel_file)
        os.remove(temp_file)
        return True
    except Exception as e:
        print(f"❌ Save Error: {str(e)} ❌")
        return False

def main():
    start_time = time.time() 
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        excel_file = get_excel_file()
        
        try:
            df = pd.read_excel(excel_file)
            if len(df.columns) < 2:
                print("❌ Need 2 columns: IP and PORT ❌")
                return
        except Exception as e:
            print(f"❌ File error: {str(e)} ❌")
            return

        if 'RESULT' not in df.columns:
            df['RESULT'] = ''

        for index, row in df.iterrows():
            ip = str(row.iloc[0])
            port = str(row.iloc[1])
            os.system('cls' if os.name == 'nt' else 'clear')
            print(f"\nTesting {ip}:{port}")
            
            ps_command = f"Test-NetConnection -ComputerName {ip} -Port {port}"
            process = subprocess.Popen(
                ['powershell', ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

            output = []
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    print(f"  {line.strip()}")
                    output.append(line.strip())

            df.at[index, 'RESULT'] = '\n'.join(output)

        if save_data(df, excel_file):
            os.system('cls' if os.name == 'nt' else 'clear')
            print(f"\n✅ Saved results to: {excel_file} ✅")
        else:
            print("\n❌ Save failed. Check file permissions. ❌")
    finally:
        elapsed_time = time.time() - start_time
        print(f"\nTime taken: {elapsed_time:.2f} seconds")

main()