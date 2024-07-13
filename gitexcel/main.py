import subprocess
import os
import json
from shutil import rmtree
from . import helpers
from . import py_git
from . import log
from .excel_convert import *


ALL_EXCEL_AS_TEXT_PATH='EXCEL_TEXT'
ALL_EXCEL_PATH='EXCEL'
METADATA_FILE_PATH='EXCEL_METADATA.json'
CHANGES_LOG_PATH='CHANGES.log'
DEBUG_LOG_PATH='DEBUG.log'

step_time=0.0

def init(path):
    subprocess.run(f'git init {path}')
    apply(path)


def clone(path, url):
    path=path.strip('/')
    project_path=f'{path}/{helpers.get_filename_or_lastfoldername(url)}'
    subprocess.run(f'git clone {url} {project_path}')
    apply(project_path)


def apply(path):
    path=path.strip('/')
    print(f'APPLY CONFIG FOR PROJECT DIR: {path}')
    os.makedirs(f"{path}/{ALL_EXCEL_PATH}", exist_ok=True)
    os.makedirs(f"{path}/{ALL_EXCEL_AS_TEXT_PATH}", exist_ok=True)
    
    print('Create pre-commit')
    hook_commands="""#!/bin/sh\npip install gitexcel\ngitexcel auto_convert --path .\nEXIT_CODE=$?\nexit $EXIT_CODE"""
    # with open(f"{path}/.githooks/pre-commit", 'w') as wf:
    with open(f"{path}/.git/hooks/pre-commit", 'w') as wf:
        wf.write(hook_commands)
    # subprocess.run(f'git config core.hooksPath {path}/.githooks', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    print('Create Excels metadata')
    excel_metadata={
        "ExcelName": {
            "SheetName": {
                "KEYS": ["ColName"], 
                "HEADER_LINE": 0}
        }
    }
    with open(f"{path}/{METADATA_FILE_PATH}", 'w') as wf:
        json.dump(excel_metadata, wf, indent=2)

    # git_attributes="*.xlsx merge=ours"
    # with open(f"{path}/.gitattributes", 'w') as wf:
    #     wf.write(git_attributes)

    git_ignore="~$*.xlsx"
    with open(f"{path}/.gitignore", 'w') as wf:
        wf.write(git_ignore)

    print('Create changes log file')
    changes_log="Excel changes will be displayed here (no update to the repo)"
    with open(f"{path}/{CHANGES_LOG_PATH}", 'w') as wf:
        wf.write(changes_log)


def convert(action, run_excel_path):
    #Find excel paths need to run
    log.init(DEBUG_LOG_PATH)
    run_excel_paths=[]
    unused_excel_text_paths=[]
    excel_paths_to_parse=[]
    excel_paths_to_gen=[]

    if run_excel_path!='.': # improve later
        run_excel_paths=[run_excel_path]
    else:
        #auto convert mode
        #get all excel files in folder
        all_excel_names=[]
        all_excel_paths=[]
        for root, dirs, files in os.walk(ALL_EXCEL_PATH):
            excel_per_folder=[f for f in files if not f.startswith('~$') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
            all_excel_names.extend([helpers.get_filename_or_lastfoldername(n) for n in excel_per_folder])
            all_excel_paths.extend([os.path.join(root, f) for f in excel_per_folder])
        dup=helpers.find_duplicates(all_excel_names)
        if dup:
            msg="\n"
            for filename, pos in dup.items():
                msg+=f"{filename}: {[all_excel_paths[k] for k in pos]}"+'\n'
            raise ValueError(f'Found duplicates excel file name: {msg}')
        
        run_excel_paths=all_excel_paths
        
        #remove unused folder
        all_excel_text_paths=[ALL_EXCEL_AS_TEXT_PATH+'/'+path for path in os.listdir(ALL_EXCEL_AS_TEXT_PATH) 
                            if os.path.isdir(os.path.join(ALL_EXCEL_AS_TEXT_PATH, path))]
        unused_excel_text_paths=[path for path in all_excel_text_paths if helpers.get_filename_or_lastfoldername(path) not in all_excel_names]

    for excel_path in run_excel_paths:
        new, upd, _ = py_git.get_changed_files(excel_path)
        if upd or new:
            excel_paths_to_parse.append(excel_path)
        else:
            text_path=f"{ALL_EXCEL_AS_TEXT_PATH}/{helpers.get_filename_or_lastfoldername(excel_path)}"
            if os.path.exists(text_path):
                new, upd, rmv = py_git.get_changed_files(text_path)
                if new or upd or rmv:
                    excel_paths_to_gen.append(excel_path)
            else:
                excel_paths_to_parse.append(excel_path)

    log.print_log_info('-> Prepare to convert: ')
    log.print_log_info(f"Excel file requested:\n{helpers.list_str(run_excel_paths)}")
    log.print_log_info(f"Excel text will be removed:\n{helpers.list_str(unused_excel_text_paths)}")
    log.print_log_info(f"Excel file -> text file:\n{helpers.list_str(excel_paths_to_parse)}")
    log.print_log_info(f"Excel file <- text file:\n{helpers.list_str(excel_paths_to_gen)}")

    for path in unused_excel_text_paths:
        rmtree(path)

    log.print_log_info('-> Convert:')
    if excel_paths_to_gen or excel_paths_to_parse:
        ALPHABET_COL_NAME=helpers.column_name_generator()   
        with open(METADATA_FILE_PATH, 'r') as rf:
            CF=json.load(rf)
    
        if excel_paths_to_parse and action!='to_excel':
            log.print_log_info('EXCEL -> TEXT')
            log.print_log_info('-------------')
            for path in excel_paths_to_parse:
                excel_name=helpers.get_filename_or_lastfoldername(path)
                log.print_log_info(f"{path} -> {ALL_EXCEL_AS_TEXT_PATH}/{excel_name}")
                excel_cf={}
                if excel_name in CF.keys(): 
                    excel_cf=CF[excel_name]
                try:
                    rmtree(f"{ALL_EXCEL_AS_TEXT_PATH}/{excel_name}")
                except:
                    pass
                parsing_excel.gen_excel_as_text(path, f"{ALL_EXCEL_AS_TEXT_PATH}/{excel_name}", excel_cf, ALPHABET_COL_NAME)


        if excel_paths_to_gen and action!='to_text': 
            log.print_log_info('TEXT -> EXCEL')
            log.print_log_info('-------------')  
            for path in excel_paths_to_gen:
                excel_name=helpers.get_filename_or_lastfoldername(path)
                text_path=f"{ALL_EXCEL_AS_TEXT_PATH}/{excel_name}"
                log.print_log_info(f"{text_path} -> {path}")
                generate_excel.gen_excel_from_text(path, text_path, ALPHABET_COL_NAME)
                subprocess.run(f'git add {ALL_EXCEL_PATH}')
    
    log.print_log_info("-> Modify commit")
    subprocess.run(f'git add {ALL_EXCEL_PATH}')
    subprocess.run(f'git add {ALL_EXCEL_AS_TEXT_PATH}')
    log.create_changes_log(CHANGES_LOG_PATH, py_git.get_changed_files(ALL_EXCEL_AS_TEXT_PATH))
    subprocess.run(f'git restore --staged {CHANGES_LOG_PATH}')
    subprocess.run(f'git restore --staged {DEBUG_LOG_PATH}')
    log.print_log_info("-> Finished")




