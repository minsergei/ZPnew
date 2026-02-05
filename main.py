import tkinter as tk
import shutil, os
from tkinter import filedialog, messagebox, scrolledtext
from create_xls import create_zp
from send_mail import mail_for_employees


def select_file_for_calculation():
    # Открываем диалог выбора файла
    file_path = filedialog.askopenfilename(
        title="Выберите файл для обработки",
        filetypes=(("Excel files", "*.xls"), ("All files", "*.*"))
    )
    # Если файл выбран, записываем путь в поле ввода
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)


def select_file_with_employeers():
    # Открываем диалог выбора файла
    file_path2 = filedialog.askopenfilename(
        title="Выберите файл для обработки",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    # Если файл выбран, записываем путь в поле ввода
    if file_path2:
        entry_path2.delete(0, tk.END)
        entry_path2.insert(0, file_path2)


def execute_process():
    path = entry_path.get()
    if not path:
        messagebox.showwarning("Внимание", "Сначала выберите файл!")
        return
    # функция обработки и создание расчеток сотрудникам
    try:
        create_zp(path)
        messagebox.showinfo("Успех", f"Файл {path} успешно обработан. Созданы расчетные листы.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")


def sending_process():
    path2 = entry_path2.get()
    if not path2:
        messagebox.showwarning("Внимание", "Сначала выберите файл!")
        return
    # функция  отправки расчеток сотрудникам
    try:
        output_field.delete(1.0, tk.END)  # Очистка поля перед выводом
        for i in mail_for_employees(path2):
            output_field.insert(tk.INSERT, f"{i}\n")
        messagebox.showinfo("Успех", "Расчетные листы отправлены по почте.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")


def delete_files():
    try:
        shutil.rmtree('calculations/')
        os.mkdir('calculations/')
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")


# Создание основного окна
root = tk.Tk()
root.title("Расчетные документы")
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
# Вычисляем координаты окна приложения
window_width = 500
window_height = 300
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Элементы интерфейса
label = tk.Label(root, text="Выберите файл с расчетными листами:")
label.pack(pady=5)

# Фрейм для размещения строки пути и кнопки обзора
frame = tk.Frame(root)
frame.pack(padx=10, fill='x')

entry_path = tk.Entry(frame)
entry_path.pack(side=tk.LEFT, expand=True, fill='x', padx=(0, 5))

btn_browse = tk.Button(frame, text="Обзор...", command=select_file_for_calculation)
btn_browse.pack(side=tk.RIGHT)


label = tk.Label(root, text="Выберите файл для каких сотрудников отправить:")
label.pack(pady=5)

# Фрейм для размещения строки пути и кнопки обзора
frame = tk.Frame(root)
frame.pack(padx=10, fill='x')

entry_path2 = tk.Entry(frame)
entry_path2.pack(side=tk.LEFT, expand=True, fill='x', padx=(0, 5))

btn_browse2 = tk.Button(frame, text="Обзор...", command=select_file_with_employeers)
btn_browse2.pack(side=tk.RIGHT)


# Кнопки выполнения основной функции
row_frame = tk.Frame(root)
row_frame.pack(pady=10)

btn1 = tk.Button(row_frame, text="Сформировать файлы", command=execute_process, width=18, bg="green", fg="white", font=('Arial', 10, 'bold'))
btn1.pack(side=tk.LEFT, padx=5)

btn2 = tk.Button(row_frame, text="Отправить файлы", command=sending_process, width=18, bg="green", fg="white", font=('Arial', 10, 'bold'))
btn2.pack(side=tk.LEFT, padx=5)



btn_run = tk.Button(root, text="Удалить созданные файлы", command=delete_files, bg="green", fg="white", font=('Arial', 10, 'bold'))
btn_run.pack(pady=5)

output_field = scrolledtext.ScrolledText(root, width=60, height=5)
output_field.pack(padx=10, pady=10)

root.mainloop()