import pandas as pd
import os
import json
import openpyxl
import shutil
import time
import re
import traceback

ALL_EXCEL_AS_TEXT_PATH='./excel_as_text/'
ALL_EXCEL_PATH='./excel/'
METADATA_FILE_PATH='./EXCEL_METADATA.json'
step_time=0

def parsing_format(x):
    if hasattr(x, '__dict__'):
        res = {'objectType': str(type(x).__name__), 'measure': {}}
        res_name=[]
        for k, v in vars(x).items():
            if hasattr(v, '__dict__'):
                format_name, format=parsing_format(v)
                if format['measure']: 
                    res['measure'][k] = format
                    res_name.append(format_name)
            else:
                if v:
                    res['measure'][k] = v 
                    res_name.append(f"{k}={str(v)}")
        return '  '.join(res_name), res
    else:
        return x, {'objectType': str(type(x).__name__), 'measure': x}
    

def log_with_timer(msg):
    global step_time
    current_time=time.time()
    if step_time==0: step_time=current_time
    print(f'(after {"{:.3f}".format(current_time-step_time)}) {msg}')
    step_time=current_time


def gen_excel_as_text(excel_path, text_path, excel_cf, ALPHABET_COL_NAME):
    workbook = openpyxl.load_workbook(excel_path)

    try:
        shutil.rmtree(text_path)
    except:
        pass

    for sheet_name, sheet_properties in excel_cf.items():
        try:
            sheet_keys=sheet_properties['KEYS']
        except:
            sheet_keys=False

        try:
            sheet_header_line=sheet_properties['HEADER_LINE']
        except:
            sheet_header_line=False


        #GET DATA
        log_with_timer(f"{sheet_name}: reading data...")
        try:
            df_values=pd.read_excel(excel_path, sheet_name=sheet_name, index_col=False, header=None)
        except ValueError:
            print(f'Not found excel file [{sheet_name}] or sheet {sheet_name}')
            return None
        df_values=df_values.rename(columns=dict(zip(list(df_values.columns), ALPHABET_COL_NAME[:len(list(df_values.columns))])))


        #GET LAYOUT
        log_with_timer(f"{sheet_name}: reading style...")
        sheet=workbook[sheet_name]
        layout_data = []
        general_format={'width': {}, 'font':{}, 'border':{}, 'fill':{}, 'number_format': {}, 'protection':{}, 'alignment':{}}
        for col_name in ALPHABET_COL_NAME[:len(list(df_values.columns))]:
            general_format['width'][col_name] = sheet.column_dimensions[col_name].width
        for row in sheet.iter_rows():
            row_layout_data = []
            # start_time = time.time()  # Start the timer
            for cell in row:
                cell_layout_data = {
                    'font': parsing_format(cell.font),
                    'border': parsing_format(cell.border),
                    'fill': parsing_format(cell.fill),
                    'number_format': parsing_format(cell.number_format),
                    'protection': parsing_format(cell.protection),
                    'alignment': parsing_format(cell.alignment)
                }

                for format_type, name_format in cell_layout_data.items():
                    name, format = name_format
                    cell_layout_data[format_type]=name
                    format['objectType']=format_type
                    if name not in general_format[format_type].keys():
                        general_format[format_type][name]=format

                row_layout_data.append(cell_layout_data)
            layout_data.append(row_layout_data)
        

        #WRITE TO READABLE FILES
        log_with_timer(f"{sheet_name}: writing data...")
        for i in range(len(df_values)):
            line=i+1
            record_title=""
            if sheet_header_line and line == sheet_header_line:
                record_title='HEADER'
            elif line > sheet_header_line:
                if sheet_header_line:
                    df_values=df_values.rename(columns=dict(zip(list(df_values.columns), df_values.iloc[sheet_header_line-1])))
                    sheet_header_line=False
                if sheet_keys:
                    record_keys=dict(zip(sheet_keys, [df_values[key].iloc[i] for key in sheet_keys]))
                    record_title=' & '.join([re.sub(r'\W+','', str(v)) for k, v in record_keys.items()])

            if len(record_title)>0:
                record_name=f'L{line}_ {record_title}'
            else:
                record_name=f'L{line}'

            record_path=f"{text_path}{sheet_name}/{record_name}/"
            record_data_path=f"{record_path}values.csv"   
            record_layout_path=f"{record_path}styles.json"

            os.makedirs(record_path)
            df_values.iloc[i].T.to_csv(record_data_path)
            
            with open(record_layout_path, 'w') as wf:
                json.dump(dict(zip(df_values.columns, layout_data[i])), wf, indent=2)
        
        try:
            with open(f"{text_path}{sheet_name}/styles_detail.json", 'w') as wf:
                    json.dump(general_format, wf, indent=2)
        except FileNotFoundError:
            pass

    
    