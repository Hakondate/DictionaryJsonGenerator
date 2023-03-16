import openpyxl
import json
from read_sheet import read_sheet

INPUT_FILE_PATH = 'data/'
INPUT_FILE_NAME = 'table_data.xlsx'
SHEET_NAME = '表全体'
OUTPUT_PATH = 'output/'
OUTPUT_FILE_NAME = 'dictionary.json'

def fetch_sheet():
    load_book = openpyxl.load_workbook(INPUT_FILE_PATH + INPUT_FILE_NAME)
    return load_book[SHEET_NAME]

def write_json(dictionary:dict):
    with open(OUTPUT_PATH + OUTPUT_FILE_NAME, mode = 'w', encoding = 'utf-8') as file:
        file.write(json.dumps(dictionary, ensure_ascii = False, indent = 4))

def main():
    sheet = fetch_sheet()
    dictionary = read_sheet(sheet)
    write_json(dictionary)

if __name__ == '__main__':
    main()
