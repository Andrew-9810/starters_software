import datetime
import yaml

import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import ttk


SAVE_FILE_CSV = ""
SAVE_FILE_YAML = ""
SHOW_TXT = ""
SHOW_CSV = ""


def c_wo(open_path, save_path):
    with open(open_path) as f:
        templates = yaml.safe_load(f)

    for unit in templates['Data']:
        if unit['Runner']['StartStatus'] == 'DNS':
            bib = f'{unit["Runner"]["Bib"]}'.rjust(4)
            total_time = f'{datetime.datetime.today().strftime("%H:%M:%S")},00'
            status = '8\n'

            fin_line = f'{bib} {total_time} {status}'
            with open(save_path, 'a') as f:
                f.write(fin_line)


def wo_c(open_path, save_path):
    """Принимает экспорт csv WinOrient, возвращает csv для O Checklist"""
    with open(open_path, 'r', encoding='1251') as input_file:
        df = pd.read_csv(
            input_file, delimiter=';', encoding='1251',
            names=[
                'Name', 'Class', 'Team', 'Year', 'Bib_Number', 'Card_Num',
                'arenda', 'Start_T', 'zx_1', 'xz_2', 'xz_3', 'xz_4', 'xz_5',
                'xz_6'
            ]
        )

        df['Start_N'] = ''
        df['Start_T'] = pd.to_datetime(df['Start_T']).dt.round('S').dt.time
        df = df[[
            'Name', 'Team', 'Class', 'Bib_Number',
            'Start_T', 'Card_Num', 'Start_N'
        ]]
        df['Team'] = df['Team'].str.replace(',', '.')

        with open(save_path, 'w') as output_file:
            df.to_csv(
                output_file, sep=',', encoding='utf-8', header=False,
                index=False
            )


def open_file_csv():
    """Получает путь к экспорту .csv WinOrient,
    запускает конвертацию файла .csv WinOrient в файл .csv для O Checklist."""
    global SAVE_FILE_CSV
    filepath = filedialog.askopenfilename()
    SAVE_FILE_CSV = filedialog.asksaveasfilename()
    wo_c(filepath, SAVE_FILE_CSV)


def open_file_yaml():
    """Получает путь к .yaml файлу O Checklist,
        запускает конвертацию файла yaml O Checklistcsv
        в файл финишной таблицы .txt для O WinOrient."""
    global SAVE_FILE_YAML
    filepath = filedialog.askopenfilename()
    SAVE_FILE_YAML = filedialog.asksaveasfilename()
    c_wo(filepath, SAVE_FILE_YAML)


def main():
    """Запуск графической оболочки."""

    root = Tk()
    root.title("Электронная шахматка")
    root.geometry("600x300")

    for c in range(3): root.columnconfigure(index=c, weight=1)
    for r in range(2): root.rowconfigure(index=r, weight=1)

    table_text_wo_c = ttk.Label(root, text="WinOrient->O Checklist")
    table_text_wo_c.grid(column=0, row=0)
    help_table_text_wo_c = ttk.Label(root, text="Справка")
    help_table_text_wo_c.grid(column=1, row=0)

    table_text_c_wo = ttk.Label(root, text="O Checklist->WinOrient")
    table_text_c_wo.grid(column=2, row=0)
    help_table_text_c_wo = ttk.Label(root, text="Справка")
    help_table_text_c_wo.grid(column=3, row=0)

    download_csv = ttk.Button(
        root, text="Загрузить .csv", command=open_file_csv
    )
    download_csv.grid(column=0, row=1)
    help_download_csv = ttk.Button(root, text="?")
    help_download_csv.grid(column=1, row=1)

    show_csv = ttk.Label(root, text=f"{SAVE_FILE_CSV}")
    show_csv.grid(column=0, row=2)

    download_yaml = ttk.Button(
        root, text="Загрузить .YAML", command=open_file_yaml
    )
    download_yaml.grid(column=2, row=1)
    help_download_yaml = ttk.Button(root, text="?")
    help_download_yaml.grid(column=3, row=1)

    show_txt = ttk.Label(root, text=f"{SAVE_FILE_YAML}")
    show_txt.grid(column=2, row=2)

    root.mainloop()


if __name__ == "__main__":
    main()
