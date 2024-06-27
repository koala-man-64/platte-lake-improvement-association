import os
import argparse
import subprocess
from github import Github
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

def commit_and_push_changes(repo, output_dir, commit_message):
    contents = repo.get_contents(output_dir, ref="main")
    for content_file in contents:
        repo.delete_file(content_file.path, "cleanup old files", content_file.sha, branch="main")
    
    for root, dirs, files in os.walk(output_dir):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                git_file_path = os.path.relpath(file_path, start=output_dir)
                repo.create_file(f"{output_dir}/{git_file_path}", commit_message, content, branch="main")

if __name__ == "__main__":
    file_path = os.getenv('FILE_PATH')
    output_dir = os.getenv('OUTPUT_DIR')
    commit_message = os.getenv('COMMIT_MESSAGE')
    github_token = os.getenv('GITHUB_TOKEN')
    extract_vba_code(file_path, output_dir)

    # Authenticate to GitHub
    g = Github(github_token)
    repo = g.get_repo(repo_name)

    # Commit and push changes
    commit_and_push_changes(repo, output_dir, commit_message)

