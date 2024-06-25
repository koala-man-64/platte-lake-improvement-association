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
    parser = argparse.ArgumentParser(description="Extract VBA macros from an Excel file and commit to GitHub.")
    parser.add_argument('file_path', help="Path to the Excel file")
    parser.add_argument('output_dir', help="Path to the output directory")
    parser.add_argument('commit_message', help="Commit message for the changes")

    args = parser.parse_args()
    file_path = args.file_path
    output_dir = args.output_dir
    commit_message = args.commit_message

    extract_vba_code(file_path, output_dir)

    # Stage the changes
    run_git_command(f'git add {output_dir}')
    
    # Commit the changes
    run_git_command(f'git commit -m "{commit_message}"')
    
    # Push the changes
    run_git_command('git push')

