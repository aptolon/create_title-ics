import tkinter as tk

root = tk.Tk()
root.title("Форма с оформлением, как в Excel")

# Настройка шрифта
font = ("Arial", 10)

# Создаем основной контейнер
main_frame = tk.Frame(root)
main_frame.pack(padx=10, pady=10)

# Функция для создания ячеек с полями ввода
def create_cell(row, col, label_text):
    # Создаем Frame для ячейки
    cell_frame = tk.Frame(main_frame, bd=1, relief="solid", padx=5, pady=5)
    cell_frame.grid(row=row, column=col, padx=2, pady=2)

    # Добавляем метку
    label = tk.Label(cell_frame, text=label_text, font=font)
    label.pack(side="left")

    # Добавляем поле для ввода
    entry = tk.Entry(cell_frame, font=font)
    entry.pack(side="right")
    
    return entry

# Создаем ячейки для ввода
entry_name = create_cell(0, 0, "Имя")
entry_age = create_cell(1, 0, "Возраст")
entry_city = create_cell(2, 0, "Город")

# Кнопка для отправки данных
def submit_form():
    name = entry_name.get()
    age = entry_age.get()
    city = entry_city.get()
    print(f"Имя: {name}, Возраст: {age}, Город: {city}")

# Кнопка для отправки данных
submit_button = tk.Button(root, text="Отправить", font=font, command=submit_form)
submit_button.pack(pady=10)

root.mainloop()
