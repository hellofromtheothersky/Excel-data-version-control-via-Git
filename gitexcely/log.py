import logging
import time


def init(log_path):
    with open(log_path, 'w') as wf:
        wf.write('')
    global g_logger
    # Create a logger
    g_logger = logging.getLogger('my_logger')
    g_logger.setLevel(logging.DEBUG)

    # Create a formatter to define the log format
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Create a file handler to write logs to a file
    file_handler = logging.FileHandler(log_path)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # # Create a stream handler to print logs to the console
    # console_handler = logging.StreamHandler()
    # console_handler.setLevel(logging.INFO)  # You can set the desired log level for console output
    # console_handler.setFormatter(formatter)

    # Add the handlers to the logger
    g_logger.addHandler(file_handler)
    # g_logger.addHandler(console_handler)

    # # Now you can log messages with different levels
    # logger.debug('This is a debug message')
    # logger.info('This is an info message')
    # logger.warning('This is a warning message')
    # logger.error('This is an error message')


def print_log_info(msg: str):
    print(msg)
    global g_logger
    g_logger.info(msg)


def log_with_timer(msg: str, last_step_time: float = None) -> float:
    current_time=time.time()
    if not last_step_time: 
        print_log_info(msg)
    else:
        print_log_info(f'(after {"{:.3f}".format(current_time-last_step_time)}) {msg}')
    return current_time


def create_changes_log(changes_log_path, changes):
    logs=dict()
    type_of_changes=['CREATED', 'UPDATED', 'REMOVED']
    for i in range(len(changes)):
        for j in range(len(changes[i])):
            try:
                _, excel_name, sheet_name, _, excel_element=changes[i][j].split('/') #EXCEL_TEXT/HiWorld/Sheet1/L1/styles.json
                excel_element=excel_element.split('.')[0]
            except:
                pass
            else:
                try:
                    logs[excel_name][sheet_name][type_of_changes[i]][excel_element]+=1
                except KeyError:
                    if excel_name not in logs.keys():
                        logs[excel_name]={}
                    if sheet_name not in logs[excel_name].keys():
                        logs[excel_name][sheet_name]={}
                    if type_of_changes[i] not in logs[excel_name][sheet_name].keys():
                        logs[excel_name][sheet_name][type_of_changes[i]]={'values': 0, 'styles': 0}
                    logs[excel_name][sheet_name][type_of_changes[i]][excel_element]+=1


    msg="Changes:\n"
    for excel_name, sheet_info in logs.items():
        msg+=f"- Excel: \"{excel_name}\":\n"
        for sheet_name, change_info in sheet_info.items():
            msg+=f"\t- Sheet: \"{sheet_name}\":\n"
            for change_type, change_detail in change_info.items():
                msg+=f"\t\t- {change_detail['values']} value rows {change_type} ({change_detail['styles']} style rows {change_type})\n"
    with open(changes_log_path, 'w') as wf:
        wf.write(msg)