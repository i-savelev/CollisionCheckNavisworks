import tkinter as tk
from tkinter import filedialog
from tkinter import Label
import webbrowser


class Interface():
    """
    Интерфейс программы
    """
    def __init__(self, func1, func2):
        self.func1 = func1
        self.func2 = func2
        self.folder_path = ''
        self.file_path = ''

    def select_file(self):
        self.file_path = filedialog.askopenfilename(title="Выберите файл")
        self.file_label.config(text=self.file_path)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory(title="Выберите папку для сохранение xml")
        self.folder_label.config(text=self.folder_path)

    def open_link_habr(self, event=None):
        webbrowser.open("https://habr.com/ru/articles/834228/")

    def open_link_github(self, event=None):
        webbrowser.open("https://github.com/i-savelev/CollisionCheckNavisworks")

    def open_link_github_current(self, event=None):
        webbrowser.open("https://github.com/barmaksis/CollisionCheckNavisworks.git")

    def set_button_function1(self):
        self.func1(self.file_path, self.file_list_box)

    def set_button_function2(self):
        self.func2(self.file_path, self.folder_path, self.file_list_box)
        
    def inicialize(self):
        self.root = tk.Tk()
        self.root.title("Создание проверок и поисковых наборов Navisworks")

        # Кнопка выбора файла excel
        self.folder_buton = tk.Button(self.root, text="Выбрать файл excel", command=self.select_file)
        self.folder_buton.grid(row=1, column=0, padx=(10, 0), pady=10)
        # Текст для передачи пути к файлу
        self.file_label = tk.Label(self.root, text="", width=80)
        self.file_label.grid(row=1, column=1, padx=(10, 10))

        # Кнопка выбора папки
        self.folder_path_button = tk.Button(self.root, text="Выбрать папку для сохранение xml", command=self.select_folder)
        self.folder_path_button.grid(row=2, column=0, padx=(10, 0), pady=10)
        # Текст для передачи пути к файлу
        self.folder_label = tk.Label(self.root, text="", width=80)
        self.folder_label.grid(row=2, column=1, padx=(10, 10))

        # Кнопка для запуска функции
        self.function1_button = tk.Button(self.root, text="Шаг 1", command=self.set_button_function1)
        self.function1_button.grid(row=3, column=0, padx=(10, 0), pady=10)

        # Кнопка для запуска функции
        self.function2_button = tk.Button(self.root, text="Шаг 2", command=self.set_button_function2)
        self.function2_button.grid(row=4, column=0, padx=(10, 0), pady=10)

        # Текстовое поле
        file_list_frame = tk.Frame(self.root)
        file_list_frame.grid(row=5, column=1, padx=(10, 0), pady=10)

        self.file_list_box = tk.Listbox(file_list_frame, height=5, width=100)
        scrollbar = tk.Scrollbar(file_list_frame,
            orient="horizontal", command=self.file_list_box.xview)
        
        self.file_list_box.config(xscrollcommand=scrollbar.set)
        self.file_list_box.grid(row=1, column=1, padx=(10, 0), pady=10, sticky="nsew")
        scrollbar.grid(row=2, column=1, padx=(10, 0), sticky="ew")

        self.habr_link  = Label(self.root, text="Инструкция", fg="blue", cursor="hand2")
        self.habr_link.grid(row=6, column=0, rowspan=2, sticky="nsew")
        self.github_link  = Label(self.root, text="Original Github repo", fg="blue", cursor="hand2")
        self.github_link.grid(row=6, column=1)
        self.github_link_current = Label(self.root, text="Current Github fork", fg="blue", cursor="hand2")
        self.github_link_current.grid(row=7, column=1)

        self.habr_link.bind("<Button-1>", self.open_link_habr)
        self.github_link.bind("<Button-1>", self.open_link_github)
        self.github_link_current.bind("<Button-1>", self.open_link_github_current)

        self.root.mainloop()