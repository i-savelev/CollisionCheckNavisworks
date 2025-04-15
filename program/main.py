"""
pyinstaller --onefile --collect-all openpyxl --name CollisionCheckNavisworks --windowed program/main.py
"""

from interface import Interface
import tkinter as tk
import utils

def step1(path_excel_file, list_box):
    list_box.delete(0, tk.END)  # очистка текстового поля
    try:
        utils.step_1(path_excel_file)
        list_box.insert(tk.END, f'Готово. Лист "Модели_Элементы" создан')
    except Exception as e:
        list_box.insert(tk.END, f'!Ошибка: {e}') # Вывод ошибок

def step2(path_excel_file, folder_xml_file, list_box):
    list_box.delete(0, tk.END)  # очистка текстового поля
    try:
        utils.step_2(path_excel_file, folder_xml_file)
        list_box.insert(tk.END, f'Готово. Созданы листы "Матрица_Заполнение" и "Списки_Заполнение"')
        list_box.insert(tk.END, f'Созданы файлы проверок и поисковых наборов "Проверки.xml" и "Поисковые наборы.xml"')
    except Exception as e:
        list_box.insert(tk.END, f'!Ошибка: {e}') # Вывод ошибок
    
interface = Interface(step1, step2)
interface.inicialize()