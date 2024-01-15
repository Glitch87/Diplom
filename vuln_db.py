import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os 
import webbrowser
import win32api

class AddVulnerabilityWindow:
    def __init__(self, parent, add_to_database_method):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Добавить уязвимость")

        # Название уязвимости
        tk.Label(self.window, text="Название уязвимости:").grid(row=0, column=0)
        self.name_entry = tk.Entry(self.window)
        self.name_entry.grid(row=0, column=1)

        # Описание уязвимости
        tk.Label(self.window, text="Описание уязвимости:").grid(row=1, column=0)
        self.description_entry = tk.Entry(self.window)
        self.description_entry.grid(row=1, column=1)

        # Ссылка
        tk.Label(self.window, text="Ссылка:").grid(row=2, column=0)
        self.link_entry = tk.Entry(self.window)
        self.link_entry.grid(row=2, column=1)

        # CVE
        tk.Label(self.window, text="CVE:").grid(row=3, column=0)
        self.cve_entry = tk.Entry(self.window)
        self.cve_entry.grid(row=3, column=1)

        # Уязвимая служба
        tk.Label(self.window, text="Уязвимая служба:").grid(row=4, column=0)
        self.service_entry = tk.Entry(self.window)
        self.service_entry.grid(row=4, column=1)

        # Кнопка "Добавить"
        tk.Button(self.window, text="Добавить", command=self.add_to_database).grid(row=5, column=0, columnspan=2, pady=10)
        self.add_to_database_method = add_to_database_method

    def add_to_database(self):
        # Получаем данные из полей ввода
        name = self.name_entry.get()
        description = self.description_entry.get()
        link = self.link_entry.get()
        cve = self.cve_entry.get()
        service = self.service_entry.get()

        # Проверяем, что все поля заполнены
        if not name or not description or not link or not cve or not service:
            messagebox.showerror("Ошибка", "Заполните все поля")
            return

        # Добавляем данные в базу данных
        self.add_to_database_method(name, description, link, cve, service)
        

        # Закрываем окно добавления уязвимости
        self.window.destroy()


class VulnerabilityDBApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vulnerability Database App")

        # Создаем базу данных или подключаемся к существующей
        self.conn = sqlite3.connect("vulnerability_db.db")
        self.cursor = self.conn.cursor()

        # Создаем таблицу, если ее нет
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS vulnerabilities
                               (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                name TEXT NOT NULL,
                                description TEXT NOT NULL,
                                link TEXT NOT NULL,
                                cve TEXT NOT NULL UNIQUE,
                                service TEXT NOT NULL)''')
        self.conn.commit()
        self.tree = ttk.Treeview(self.root, columns=("Название", "Описание", "Ссылка", "CVE", "Служба"), show="headings")
        self.tree.grid(row=1, column=0, rowspan=3, columnspan=3, pady=20)
        self.tree.bind("<ButtonRelease-1>", self.update_selected)
        # Создаем и размещаем виджеты
        self.create_widgets()

    def create_widgets(self):
        # Кнопки
        self.display_database()
        tk.Button(self.root, text="Добавить уязвимость", command=self.show_add_vulnerability_window).grid(row=1, column=4, columnspan=2, pady=10)
        tk.Button(self.root, text="Удалить", command=self.delete_selected).grid(row=2, column=4, columnspan=2, pady=10)
        tk.Button(self.root, text="Справка", command=lambda: self.show_help("help")).grid(row=3, column=4, columnspan=2, pady=10)
        tk.Button(self.root, text="Лабораторная 1", command=lambda: self.show_help("Lab1")).grid(row=4, column=0, columnspan=1, pady=10)
        tk.Button(self.root, text="Лабораторная 2", command=lambda: self.show_help("Lab2")).grid(row=4, column=1, columnspan=1, pady=10)
        tk.Button(self.root, text="Лабораторная 3", command=lambda: self.show_help("Lab3")).grid(row=4, column=2, columnspan=1, pady=10)
        self.tree.heading("Название", text="Название")
        self.tree.heading("Описание", text="Описание")
        self.tree.heading("Ссылка", text="Ссылка")
        self.tree.heading("CVE", text="CVE")
        self.tree.heading("Служба", text="Служба")
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.bind("<Double-1>", self.open_selected_field)

    def display_database(self):
        # Выводим базу данных в новом окне
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Получаем данные из базы данных
        self.cursor.execute("SELECT name, description, link, cve, service FROM vulnerabilities")
        data = self.cursor.fetchall()

        # Выводим данные в Treeview
        for row in data:
            self.tree.insert("", "end", values=row)

    def open_selected_field(self, event):
        # Получаем выделенный элемент
        column = self.tree.identify_column(event.x)

        # Получите имя столбца
        column_name = self.tree.heading(column, "text")

        # Получаем выделенный элемент
        selected_item = self.tree.selection()

        if selected_item:
            # Получаем данные выделенного элемента
            item_data = self.tree.item(selected_item, "values")

            # В зависимости от имени столбца вызываем соответствующую функцию
            if column_name == "Название":
                self.show_selected_field_window(item_data[0])
            elif column_name == "Описание":
                self.show_selected_field_window(item_data[1])
            elif column_name == "Ссылка":
                self.show_selected_field_window(item_data[2])
            elif column_name == "CVE":
                self.show_selected_cve_window(item_data[3])
            elif column_name == "Служба":
                self.show_selected_field_window(item_data[4])

    def show_selected_field_window(self, field_data):
        # Создаем окно для отображения данных выбранного элемента
        selected_field_window = tk.Toplevel(self.root)
        selected_field_window.title("Подробная информация")

        # Выводим данные в новом окне
        tk.Label(selected_field_window, text=f"Данные поля: {field_data}").pack()
            
    def add_to_database(self, name, description, link, cve, service):
        # Проверяем, что все поля заполнены
        if not name or not description or not link or not cve or not service:
            messagebox.showerror("Ошибка", "Заполните все поля")
            return
        # Добавляем данные в базу данных
        self.cursor.execute("INSERT INTO vulnerabilities (name, description, link, cve, service) VALUES (?, ?, ?, ?, ?)",
                            (name, description, link, cve, service))
        self.conn.commit()
        messagebox.showinfo("Успех", "Данные успешно добавлены в базу данных")
        self.display_database()
    def delete_selected(self):
        # Получаем ID выбранного элемента
        selected_item = self.tree.selection()
        if selected_item:
            item_id = self.tree.item(selected_item, "values")[0]

            # Удаляем элемент из базы данных
            self.cursor.execute("DELETE FROM vulnerabilities WHERE id=?", (item_id,))
            self.conn.commit()

            # Обновляем отображение
            self.display_database()

    def update_selected(self, event):
        # Получаем данные выбранного элемента при клике
        selected_item = self.tree.selection()
        if selected_item:
            item_data = self.tree.item(selected_item, "values")
            print("Выбран элемент:", item_data)

    def show_add_vulnerability_window(self):
        # Открываем окно для добавления новой уязвимости
        add_window = AddVulnerabilityWindow(self.root, self.add_to_database)


    def show_help(self, file_name):
        # Определите путь к файлу справки в том же каталоге, где находится исполняемый файл
        help_file_path = os.path.join(os.path.dirname(__file__), file_name+".chm")

        # Проверьте, существует ли файл справки
        if os.path.exists(help_file_path):
            # Откройте справочный HTML-файл в браузере по умолчанию
            webbrowser.open(help_file_path)
        else:
            error_message = f"Файл справки не найден. Расположение исполняемого файла: {os.path.abspath(__file__)}"
            messagebox.showerror("Ошибка", error_message)

if __name__ == "__main__":
    root = tk.Tk()
    app = VulnerabilityDBApp(root)
    root.mainloop()