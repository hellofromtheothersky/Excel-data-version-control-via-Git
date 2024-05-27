import subprocess
import re
import itertools
import time

def column_name_generator():
    # Generate ALPHABET_COL_NAME of lengths 1 to 5
    global ALPHABET_COL_NAME 
    ALPHABET_COL_NAME = []
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    # ALPHABET_COL_NAME of length 1
    ALPHABET_COL_NAME.extend(list(alphabet))

    # ALPHABET_COL_NAME of length 2 to 5
    for length in range(2, 3):
        ALPHABET_COL_NAME.extend([''.join(combination) for combination in itertools.product(alphabet, repeat=length)])


def log_with_timer(msg):
    global step_time
    current_time=time.time()
    if step_time==0: step_time=current_time
    print(f'(after {"{:.3f}".format(current_time-step_time)}) {msg}')
    step_time=current_time

def get_filename_or_lastfoldername(s):
    return re.findall(r'[\\/]([^\\\./]*)[\.]{0,1}[\w]*$', s)[0]

def get_changed_files(path):
    output = subprocess.check_output(['git', 'status', '--porcelain', path])

    # Decode the output from bytes to string
    output_str = output.decode('utf-8')

    # Split the output into lines
    lines = output_str.splitlines()

    # Initialize empty lists to store the file names
    new_files = []
    modified_files = []

    # Iterate through the lines and extract the file names
    for line in lines:
        line=line.strip()
        if line.startswith('A '):
            new_files.append(re.findall(r'\w+\s+["]*(.*)["]*', line)[0])
        elif line.startswith('M '):
            modified_files.append(re.findall(r'\w+\s+["]*(.*)["]*', line)[0])
    
    return new_files, modified_files