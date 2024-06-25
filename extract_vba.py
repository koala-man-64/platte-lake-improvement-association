import os
from oletools.olevba import VBA_Parser

def extract_vba_code(excel_path, output_dir):
    vbaparser = VBA_Parser(excel_path)
    if vbaparser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_all_macros():
            module_name = os.path.splitext(os.path.basename(vba_filename))[0]
            output_file = os.path.join(output_dir, f"{module_name}.vba")
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(vba_code)
    vbaparser.close()

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

if __name__ == "__main__":
    file_path = os.getenv('FILE_PATH')
    output_dir = os.getenv('OUTPUT_DIR', 'output_directory')  # Default path to the directory
    
    # Extract VBA code
    extract_vba_code(file_path, output_dir)
