import subprocess
import re


def get_changed_files(path: str) -> tuple[list, list]:
    output = subprocess.check_output(['git', 'status', '--porcelain', path])

    # Decode the output from bytes to string
    output_str = output.decode('utf-8')

    # Split the output into lines
    lines = output_str.splitlines()

    # Initialize empty lists to store the file names
    new_files = []
    modified_files = []
    removed_files = []

    # Iterate through the lines and extract the file names
    for line in lines:
        line=line.strip()
        if line.startswith('A') or line.startswith('??'):
            new_files.append(re.findall(r'[\w?]+\s+["]*(.+?)["]*$', line)[0])
        elif line.startswith('M'):
            modified_files.append(re.findall(r'\w+\s+["]*(.+?)["]*$', line)[0])
        elif line.startswith('D'):
            removed_files.append(re.findall(r'\w+\s+["]*(.+?)["]*$', line)[0])
    
    return new_files, modified_files, removed_files
