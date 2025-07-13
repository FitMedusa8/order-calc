import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta

# --- Глобальные переменные ---
sales_df = None
weights_df = None
result_df = None

# --- Загрузка Excel-файла с продажами ---
def load_sales_file():
    global sales_df
    file_path = filedialog.askopenfilename(filetypes=[("Excel файлы", "*.xlsx")])
    if file_path:
        try:
            temp_df = pd.read_excel(file_path)
            if len(temp_df.columns) < 3:
                raise ValueError("Файл должен содержать минимум 3 колонки: SKU, Наименование и даты продаж")
            if temp_df.empty:
                raise ValueError("Файл пустой")
            sales_df = temp_df
            messagebox.showinfo("Успех", "Файл продаж загружен успешно.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {e}")

# --- Загрузка справочника веса ---
def load_weights_file():
    global weights_df
    file_path = filedialog.askopenfilename(filetypes=[("Excel файлы", "*.xlsx")])
    if file_path:
        try:
            temp_df = pd.read_excel(file_path)
            if 'SKU' not in temp_df.columns or 'Вес' not in temp_df.columns:
                raise ValueError("Справочник должен содержать колонки 'SKU' и 'Вес'")
            if temp_df.empty:
                raise ValueError("Справочник пустой")
            weights_df = temp_df[['SKU', 'Вес']]
            messagebox.showinfo("Успех", "Справочник веса загружен.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить справочник: {e}")

# --- Расчёт среднего заказа за период ---
def calculate_order():
    global result_df
    if sales_df is None or weights_df is None:
        messagebox.showwarning("Внимание", "Сначала загрузите файл продаж и справочник веса.")
        return

    try:
        period = int(period_var.get())
        data = sales_df.copy()
        data.columns = data.columns.astype(str)

        sku_col = data.columns[0]
        name_col = data.columns[1]
        date_cols = data.columns[2:]
        recent_cols = date_cols[-period:]

        weight_map = dict(zip(weights_df['SKU'].astype(str), weights_df['Вес']))
        data['Вес'] = data[sku_col].astype(str).map(weight_map)

        for col in recent_cols:
            data[col] = data.apply(
                lambda row: row[col] * row['Вес'] if row['Вес'] < 1 else row[col],
                axis=1
            )

        data['Средняя продажа'] = data[recent_cols].mean(axis=1)
        data['Рекоменд. заказ'] = data['Средняя продажа'].round(2)
        result_df = data[[sku_col, name_col, 'Рекоменд. заказ']]
        result_df.to_excel("temp_result.xlsx", index=False)
        display_result()

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка расчёта: {e}")

# --- Функция фильтрации таблицы по поисковому запросу ---
def filter_table(*args):
    query = search_var.get().lower()
    for row in tree.get_children():
        tree.delete(row)
    if result_df is None:
        return
    for idx, row in result_df.iterrows():
        sku = str(row[0]).lower()
        name = str(row[1]).lower()
        qty = row[2]
        if query in sku or query in name:
            tag = ""
            if qty < 10:
                tag = "low"
            elif qty > 250:
                tag = "high"
            tree.insert("", tk.END, values=(row[0], row[1], round(qty, 2)), tags=(tag,))

# --- Отображение результата в таблице с цветами ---
def display_result():
    filter_table()

# --- Редактирование по двойному клику ---
def on_double_click(event):
    item_id = tree.identify_row(event.y)
    column = tree.identify_column(event.x)

    if column == '#3':  # Только "Рекоменд. заказ"
        current_value = tree.item(item_id)['values'][2]
        new_value = simpledialog.askfloat("Редактирование", "Введите новое значение:", initialvalue=current_value)

        if new_value is not None:
            tree.set(item_id, column="#3", value=round(new_value, 2))
            index = tree.index(item_id)
            result_df.at[index, 'Рекоменд. заказ'] = round(new_value, 2)
            result_df.to_excel("temp_result.xlsx", index=False)

# --- Выгрузка заказа на 14 дней вперёд ---
def export_order():
    if result_df is None:
        messagebox.showwarning("Внимание", "Сначала рассчитайте заказ.")
        return

    date_str = simpledialog.askstring("Дата начала", "Введите дату начала заказа (ДД.ММ.ГГГГ):")
    try:
        start_date = datetime.strptime(date_str, "%d.%m.%Y")
    except:
        messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГГГ.")
        return

    sku_col = result_df.columns[0]
    name_col = result_df.columns[1]
    qty_col = result_df.columns[2]

    export_df = result_df[[sku_col, name_col]].copy()
    for i in range(14):
        day = (start_date + timedelta(days=i)).strftime("%d.%m.%Y")
        export_df[day] = result_df[qty_col]

    export_df.to_excel("заказ_на_14_дней.xlsx", index=False)
    messagebox.showinfo("Готово", "Файл 'заказ_на_14_дней.xlsx' сохранён.")

# --- Интерфейс ---
root = tk.Tk()
root.title("Расчёт заказа на производство")
root.geometry("800x600")

# --- PATCH: создаём search_var после root ---
search_var = tk.StringVar()  # --- PATCH: переменная для поиска ---

tk.Button(root, text="Загрузить продажи", command=load_sales_file).pack(pady=5)
tk.Button(root, text="Загрузить справочник веса", command=load_weights_file).pack(pady=5)

tk.Label(root, text="Поиск (SKU или наименование):").pack()
search_entry = tk.Entry(root, textvariable=search_var)
search_entry.pack(pady=5)
search_var.trace_add("write", filter_table)  # связываем ввод с фильтрацией

tk.Label(root, text="Период анализа (дней):").pack()
period_var = tk.StringVar(value="14")
ttk.Combobox(root, textvariable=period_var, values=["7", "14", "30"], state="readonly").pack(pady=5)

tk.Button(root, text="Рассчитать заказ", command=calculate_order).pack(pady=10)
tk.Button(root, text="Выгрузить заказ на 14 дней", command=export_order).pack(pady=5)

tree = ttk.Treeview(root, columns=("SKU", "Наименование", "Рекоменд. заказ"), show="headings")
tree.heading("SKU", text="SKU")
tree.heading("Наименование", text="Наименование")
tree.heading("Рекоменд. заказ", text="Рекоменд. заказ")
tree.pack(expand=True, fill="both", padx=10, pady=10)

tree.tag_configure("low", background="#ffcccc")
tree.tag_configure("high", background="#fff5b2")

tree.bind("<Double-1>", on_double_click)

root.mainloop()