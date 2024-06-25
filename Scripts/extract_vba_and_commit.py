import os
import argparse
import subprocess
from oletools.olevba import VBA_Parser

def extract_vba_code(excel_path, output_dir):
    vbaparser = VBA_Parser(excel_path)
    if vbaparser.detect_vba_macros():
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_all_macros():
            module_name = os.path.splitext(os.path.basename(vba_filename))[0]
            output_file = os.path.join(output_dir, f"{module_name}.vba")
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(vba_code)
    vbaparser.close()

def run_git_command(command, cwd=None):
    result = subprocess.run(command, cwd=cwd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print(result.stdout.decode())
    if result.stderr:
        print(result.stderr.decode())

if __name__ == "__main__":
    file_path = os.getenv('FILE_PATH')
    output_dir = os.getenv('OUTPUT_DIR')
    commit_message = os.getenv('OUTPUT_DIR')

    extract_vba_code(file_path, output_dir)

    # Stage the changes
    run_git_command(f'git add {output_dir}')
    
    # Commit the changes
    run_git_command(f'git commit -m "{commit_message}"')
    
    # Push the changes
    run_git_command('git push')

