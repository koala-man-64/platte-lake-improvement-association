import os
import win32com.client

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
    file_path = os.getenv('FILE_PATH')
    output_dir = os.getenv('OUTPUT_DIR', 'output_directory')  # Default path to the directory
    
    # Extract VBA code
    extract_vba_code(file_path, output_dir)
