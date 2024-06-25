import os
import requests
import win32com.client

def download_file_from_github(url, local_path):
    response = requests.get(url)
    response.raise_for_status()
    
    with open(local_path, 'wb') as f:
        f.write(response.content)

def extract_vba_code(excel_path, output_dir):
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(excel_path)
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for vb_component in workbook.VBProject.VBComponents:
        if vb_component.Type in [1, 2, 3]:  # 1 = Module, 2 = Class Module, 3 = UserForm
            code_module = vb_component.CodeModule
            code = code_module.Lines(1, code_module.CountOfLines)
            
            output_file = os.path.join(output_dir, f"{vb_component.Name}.vba")
            with open(output_file, 'w') as f:
                f.write(code)
    
    workbook.Close(SaveChanges=False)
    excel.Quit()

if __name__ == "__main__":
    github_file_url = os.getenv('GITHUB_FILE_URL')
    local_excel_path = 'downloaded_excel_file.xlsm'
    output_dir = os.getenv('OUTPUT_DIR', 'default_path_to_output_directory')
    
    # Download the Excel file from GitHub
    download_file_from_github(github_file_url, local_excel_path)
    
    # Extract VBA code
    extract_vba_code(local_excel_path, output_dir)
