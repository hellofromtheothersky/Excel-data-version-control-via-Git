import subprocess
import os
import re
import itertools
import argparse
import parsing_excel
import generate_excel
import json
import xlsxwriter
import traceback
import time


ALL_EXCEL_AS_TEXT_PATH='./excel_as_text/'
ALL_EXCEL_PATH='./excel/'
METADATA_FILE_PATH='./EXCEL_METADATA.json'

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
        if line.startswith('A '):
            new_files.append(re.findall(r'\w+\s+["]*(.*)["]*', line)[0])
        elif line.startswith('M '):
            modified_files.append(re.findall(r'\w+\s+["]*(.*)["]*', line)[0])
    
    return new_files, modified_files


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.realpath(__file__))+'/../')
    parser = argparse.ArgumentParser()
    parser.add_argument('--action',
                        default='',
                        const='',
                        nargs='?',
                        choices=['to_text', 'to_excel', ''],
                        help='define action with files')    
    parser.add_argument('--excelpath', type=str, default="")
    parser.add_argument('--wait', type=bool, default=False)
    # parser.add_argument('--exceltextpath', type=str, default="")
    args = parser.parse_args()
    

    excel_paths_to_add_auto_parsing=[]
    excel_paths_to_parse=[]
    excel_text_paths_to_gen=[]

    if not args.action:
        if not args.excelpath:
            for root, dirs, files in os.walk(ALL_EXCEL_PATH):
                files=[f for f in files if f.endswith('.xlsx') or f.endswith('.xlsm')]
                for file in files:
                    path=os.path.join(root, file)
                    new_files, updated_files = get_changed_files(path)
                    if updated_files:
                        excel_paths_to_parse.append(path)
                    elif new_files:
                        excel_paths_to_add_auto_parsing.append(path)
                        excel_paths_to_parse.append(path)
                        excel_text_paths_to_gen.append(path)

                    else:
                        text_path=ALL_EXCEL_AS_TEXT_PATH+get_filename_or_lastfoldername(file)+'/'
                        new_files, updated_files = get_changed_files(text_path)
                        if new_files or updated_files:
                            excel_text_paths_to_gen.append(path)
        elif args.excelpath:
            pass
    else:
        if args.action=='to_text' and args.excelpath:
            new_files, updated_files = get_changed_files(args.excelpath)
            if updated_files:
                excel_paths_to_parse.append(args.excelpath)
            elif new_files:
                print("No action! New file is parsing at commit time")
        elif args.action=='to_excel' and args.excelpath:
            text_path=ALL_EXCEL_AS_TEXT_PATH+get_filename_or_lastfoldername(args.excelpath)+'/'
            new_files, updated_files = get_changed_files(args.text_path)
            if new_files or updated_files:
                excel_text_paths_to_gen.append(args.excelpath)


    with open('./excel_metadata.json', 'r') as rf:
        CF=json.load(rf)
    print(excel_text_paths_to_gen)

    try:
        if excel_paths_to_parse:
            print('EXCEL -> TEXT')
            print('-------------')
            column_name_generator()       
            for path in excel_paths_to_parse:
                excel_name=get_filename_or_lastfoldername(path)
                print(f"{path} -> {ALL_EXCEL_AS_TEXT_PATH+excel_name+'/'}")
                parsing_excel.gen_excel_as_text(path, ALL_EXCEL_AS_TEXT_PATH+excel_name+'/', CF[excel_name], ALPHABET_COL_NAME)


        if excel_paths_to_add_auto_parsing:
            print('NORMAL EXCEL -> AUTO-PARSING EXCEL')
            print('-----------------------------------')
            for path in excel_paths_to_add_auto_parsing:
                newexcelpath=path.replace('.xlsx', 'xlsm')
                print(f"{path} -> {newexcelpath}")
                workbook = xlsxwriter.Workbook(newexcelpath)
                workbook.add_vba_project('AutoParsing.bin')
                workbook.close()
                

        if excel_text_paths_to_gen: 
            print('TEXT -> EXCEL')
            print('-------------')
            if not excel_paths_to_parse: 
                column_name_generator()        
            for path in excel_text_paths_to_gen:
                excel_name=get_filename_or_lastfoldername(path)
                text_path=ALL_EXCEL_AS_TEXT_PATH+excel_name+'/'
                print(f"{text_path} -> {path}")
                generate_excel.gen_excel_from_text(path, text_path, ALPHABET_COL_NAME)
    except Exception:
        traceback.print_exc()
        print("Error! Close after 30 seconds")
        if args.wait: time.sleep(30)
    else:
        print("Succeed! Close after 2 seconds")
        if args.wait: time.sleep(2)