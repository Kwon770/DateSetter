import os
from os.path import isfile, join
import sys
import re
from shutil import copyfile
from datetime import date

PROGRAM_ANNOUNCEMENT_STRING = """

______      _          _____      _   _            
|  _  \    | |        /  ___|    | | | |           
| | | |__ _| |_ ___   \ `--.  ___| |_| |_ ___ _ __ 
| | | / _` | __/ _ \   `--. \/ _ \ __| __/ _ \ '__|
| |/ / (_| | ||  __/  /\__/ /  __/ |_| ||  __/ |   
|___/ \__,_|\__\___|  \____/ \___|\__|\__\___|_|
[ 날짜 입력기 ]

날짜를 입력할 파일 목록 입니다."""
EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xlsb", "xltx", "xls", "xlt", "xml", "xlam", "xlw", "xlr"]

excel_files_len = 0

def clear_console():
    os.system('cls')
    os.system('clear')

def print_program_announcement():
    global PROGRAM_ANNOUNCEMENT_STRING
    print(PROGRAM_ANNOUNCEMENT_STRING)


def scan_directory():
    global EXCEL_EXTENSIONS
    global excel_files_len

    # current_path = sys.path[0]
    current_path = '/'.join(sys.executable.split('/')[:-1])
    excel_file_list = []
    for f in os.listdir(current_path):
        if not isfile(join(current_path, f)):
            continue

        file_info = f.split('.')
        if len(file_info) != 2 or file_info[1] not in EXCEL_EXTENSIONS:
            continue

        file_info2 = re.split('[-.]', f)
        if len(file_info2) >= 3 and file_info2[1].isnumeric() and file_info2[2].isnumeric():
            continue

        excel_file_list.append(f)

    excel_files_len = len(excel_file_list)
    if excel_files_len == 0:
        print("[시스템] 변경할 파일이 없습니다.")
        sys.exit(0)

    return excel_file_list



def print_files_list(files_list):
    for idx, file in enumerate(files_list):
        print(f'- {idx}. {file}')
    print()



def input_target_index():
    global excel_files_len

    while True:
        try:
            target_index = int(input("[입력] 변경할 파일의 번호를 입력하세요 : "))
            if target_index >= excel_files_len:
                raise ValueError

            return target_index

        except ValueError:
            print("[에러] 올바르지 않은 입력 형태입니다.")


def set_date_on_file(files_list, target_index):
    target_file = files_list[target_index]
    target_file_name, target_file_extension = target_file.split('.')
    copyfile(target_file, "TEMP")

    date_string = date.today().strftime("%y-%m-%d")
    os.rename("TEMP", target_file_name + date_string + "." + target_file_extension)


if __name__ == '__main__':
    print_program_announcement()

    files_list = scan_directory()
    print_files_list(files_list)

    target_index = input_target_index()
    set_date_on_file(files_list, target_index)