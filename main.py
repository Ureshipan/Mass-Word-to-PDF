import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import time
import threading
import tkinter.font as font
from os import listdir
from os.path import isfile, join

import sys
import os
import comtypes.client
import win32com.client as win32


wdFormatPDF = 17


def choose_input_folder():
    input_folder_path = filedialog.askdirectory(initialdir="D:/Yandex disc/YandexDisk/YaProjs/ConvertPy/test_in")
    if input_folder_path:
        input_folder_var.set(input_folder_path)


def choose_output_folder():
    output_folder_path = filedialog.askdirectory(initialdir="D:/Yandex disc/YandexDisk/YaProjs/ConvertPy/test_out")
    if output_folder_path:
        output_folder_var.set(output_folder_path)


def convert():
    inpat = input_folder_var.get().replace('/', '\\')
    outpat = output_folder_var.get().replace('/', '\\')
    if inpat != "Выберите папку с \nWord документами" and outpat != "Выберите папку \nдля PDF файлов":
        files = [f for f in listdir(inpat) if isfile(join(inpat, f)) and ('.doc' in f or '.docx' in f) and '~$' not in f]
        print(files)
        i = 0
        word = comtypes.client.CreateObject('Word.Application')
        progress_bar["maximum"] = len(files)
        for file in files:
            try:
                print(inpat + '\\' + file)
                doc = word.Documents.Open(inpat + '\\' + file, ReadOnly=True)
                outfn = outpat + '\\' + '.'.join(file.split('.')[:-1]) + '.pdf'
                print(outfn)
                doc.SaveAs(outfn, FileFormat=wdFormatPDF)
                doc.Close()
            except Exception as e:
                print(e)
            i += 1
            progress_bar["value"] = i
            root.update_idletasks()

    # Здесь вы можете добавить реальную логику конвертации


# Создаем графический интерфейс
root = tk.Tk()
root.title("Word to PDF")
root.columnconfigure(0, minsize=350)  # Колонка 0 будет иметь минимальную ширину 100
root.columnconfigure(1, minsize=350)
root.rowconfigure([0, 1, 2], minsize=80)

custom_font = font.Font(family="Helvetica", size=12, weight="bold")

input_folder_var = tk.StringVar()
input_folder_var.set("Выберите папку с \nWord документами")

output_folder_var = tk.StringVar()
output_folder_var.set("Выберите папку \nдля PDF файлов")

input_label = tk.Label(root, textvariable=input_folder_var, padx=10, pady=10, font=custom_font)
input_label.grid(row=0, column=0)

input_button = tk.Button(root, text="Выбрать папку для\n входящих файлов", command=choose_input_folder,
                         font=custom_font, bg="#DBE2EF", width=20)
input_button.grid(row=1, column=0)

output_label = tk.Label(root, textvariable=output_folder_var, padx=10, pady=10, font=custom_font)
output_label.grid(row=0, column=1)

output_button = tk.Button(root, text="Выбрать папку для\n выходящих файлов", command=choose_output_folder,
                          font=custom_font, bg="#F67280", width=20)
output_button.grid(row=1, column=1)

start_button = tk.Button(root, text="Начать конвертацию", command=convert, font=custom_font, bg="#A8E6CF")
start_button.grid(row=2, column=0, columnspan=2)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
progress_bar.grid(row=3, column=0, columnspan=2)

root.mainloop()
