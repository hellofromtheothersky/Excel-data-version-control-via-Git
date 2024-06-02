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
import py_git 
import helpers


ALL_EXCEL_AS_TEXT_PATH='./excel_as_text/'
ALL_EXCEL_PATH='./excel/'
METADATA_FILE_PATH='./EXCEL_METADATA.json'
SRC_PATH='./.src/'

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


def log_with_timer(msg:str):
    global step_time
    current_time=time.time()
    if step_time==0: step_time=current_time
    print(f'(after {"{:.3f}".format(current_time-step_time)}) {msg}')
    step_time=current_time


def get_filename_or_lastfoldername(s: str) -> str:
    return re.findall(r'[\\/]([^\\\./]*)[\.]{0,1}[\w]*$', s)[0]

import sys

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

    all_excelpaths=[]
    all_excelfiles=[]
    for root, dirs, files in os.walk(ALL_EXCEL_PATH):
        excelfiles_per_folder=[f for f in files if not f.startswith('~$') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
        all_excelfiles.extend(excelfiles_per_folder)
        all_excelpaths.extend([os.path.join(root, f) for f in excelfiles_per_folder])
    
    try:
        dup=helpers.find_duplicates(all_excelfiles)
        if dup:
            msg="\n"
            for filename, pos in dup.items():
                msg+=f"{filename}: {[all_excelpaths[k] for k in pos]}"+'\n'
            raise ValueError(f'Found duplicates excel file name: {msg}')
        
        excelpaths=[]
        if args.excelpath: 
            excelpaths.append(args.excelpath)
        else:
            excelpaths=all_excelpaths
        
        for excelpath in excelpaths:
            new, upd = py_git.get_changed_files(excelpath)
            if upd:
                excel_paths_to_parse.append(excelpath)
            elif new:
                excel_paths_to_parse.append(excelpath)
                excel_paths_to_add_auto_parsing.append(excelpath)
            else:
                print('---')
                text_path=ALL_EXCEL_AS_TEXT_PATH+get_filename_or_lastfoldername(excelpath)+'/'
                new, upd = py_git.get_changed_files(text_path)
                if new or upd:
                    excel_text_paths_to_gen.append(excelpath)

        if excel_text_paths_to_gen or excel_paths_to_parse:
            with open('./excel_metadata.json', 'r') as rf:
                CF=json.load(rf)

        if excel_paths_to_parse and args.action!='to_excel':
            print('EXCEL -> TEXT')
            print('-------------')
            column_name_generator()       
            for path in excel_paths_to_parse:
                excel_name=get_filename_or_lastfoldername(path)
                print(f"{path} -> {ALL_EXCEL_AS_TEXT_PATH+excel_name+'/'}")
                excel_cf={}
                if excel_name in CF.keys(): 
                    excel_cf=CF[excel_name]

                parsing_excel.gen_excel_as_text(path, ALL_EXCEL_AS_TEXT_PATH+excel_name+'/', excel_cf, ALPHABET_COL_NAME)
                if not args.action:
                    subprocess.run('git add .', capture_output=False)

        if excel_paths_to_add_auto_parsing and not args.action:
            print('NORMAL EXCEL -> AUTO-PARSING EXCEL')
            print('-----------------------------------')
            for path in excel_paths_to_add_auto_parsing:
                os.remove(path)
                newexcelpath=path.replace('.xlsx', '.xlsm')
                print(f"{path} -> {newexcelpath}")
                workbook = xlsxwriter.Workbook(newexcelpath)
                workbook.add_vba_project(SRC_PATH+'AutoParsing.bin')
                workbook.close()
                excel_text_paths_to_gen.append(newexcelpath)
                
        if excel_text_paths_to_gen and args.action!='to_text': 
            print('TEXT -> EXCEL')
            print('-------------')
            if not excel_paths_to_parse: 
                column_name_generator()        
            for path in excel_text_paths_to_gen:
                excel_name=get_filename_or_lastfoldername(path)
                text_path=ALL_EXCEL_AS_TEXT_PATH+excel_name+'/'
                print(f"{text_path} -> {path}")
                generate_excel.gen_excel_from_text(path, text_path, ALPHABET_COL_NAME)

        print('Checking status')
        subprocess.run('git status', capture_output=False)
        subprocess.run('git add .', capture_output=False)
    except Exception as error:
        traceback.print_exc()
        print('------')
        print(error)
        print("Error! Close after 30 seconds")
        if args.wait: time.sleep(30)
        sys.exit(1)
    else:
        print("Succeed! Close after 2 seconds")
        if args.wait: time.sleep(2)
        sys.exit(0)