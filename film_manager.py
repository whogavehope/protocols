import customtkinter as ctk
from tkinter import messagebox, ttk
import sqlite3
import pandas as pd

def get_all_films():
    conn = sqlite3.connect("films.db")
    cur = conn.cursor()
    cur.execute("SELECT sidak_num FROM films ORDER BY sidak_num")
    films = [row[0] for row in cur.fetchall()]
    conn.close()
    return films

def get_film_details(sidak):
    conn = sqlite3.connect("films.db")
    cur = conn.cursor()
    cur.execute("SELECT supplier_name, thickness FROM films WHERE sidak_num = ?", (sidak,))
    details = cur.fetchone()
    conn.close()
    return details

def add_film_window():
    def save_film():
        sidak = entry_sidak.get().strip()
        supplier = entry_supplier.get().strip()
        try:
            thickness = float(entry_thickness.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Толщина должна быть числом")
            return

        if not sidak or not supplier:
            messagebox.showerror("Ошибка", "Все поля обязательны")
            return

        conn = sqlite3.connect("films.db")
        cur = conn.cursor()
        try:
            cur.execute("INSERT INTO films (sidak_num, supplier_name, thickness) VALUES (?, ?, ?)",
                        (sidak, supplier, thickness))
            conn.commit()
            messagebox.showinfo("Успех", f"Плёнка {sidak} добавлена")
            win.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Ошибка", "Такая плёнка уже существует")
        finally:
            conn.close()

    win = ctk.CTkToplevel()
    win.title("Добавить плёнку")

    ctk.CTkLabel(win, text="Номер Sidak").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    entry_sidak = ctk.CTkEntry(win, width=300)
    entry_sidak.grid(row=0, column=1, padx=5, pady=5)

    ctk.CTkLabel(win, text="Поставщик").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    entry_supplier = ctk.CTkEntry(win, width=300)
    entry_supplier.grid(row=1, column=1, padx=5, pady=5)

    ctk.CTkLabel(win, text="Толщина").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    entry_thickness = ctk.CTkEntry(win, width=300)
    entry_thickness.grid(row=2, column=1, padx=5, pady=5)

    ctk.CTkButton(win, text="Сохранить", command=save_film, fg_color="green", hover_color="darkgreen").grid(row=3, columnspan=2, pady=10)

def delete_film_window():
    def delete_film():
        selected_film = entry_sidak.get().strip()
        if not selected_film:
            messagebox.showerror("Ошибка", "Выберите плёнку для удаления")
            return

        confirm = messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить плёнку {selected_film}?")
        if not confirm:
            return

        conn = sqlite3.connect("films.db")
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM films WHERE sidak_num = ?", (selected_film,))
            conn.commit()
            messagebox.showinfo("Успех", f"Плёнка {selected_film} удалена")
            win.destroy()
        except sqlite3.Error as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при удалении: {e}")
        finally:
            conn.close()

    win = ctk.CTkToplevel()
    win.title("Удалить плёнку")
    win.geometry("400x300")

    ctk.CTkLabel(win, text="Введите или выберите номер Sidak").pack(pady=(20, 5))

    # Создаём фильтруемое поле
    entry_sidak, results_frame = create_filterable_sidak_selector(
        parent=win,
        initial_x=50,
        initial_y=60,
        width=300,
        on_select_callback=None  # не нужно ничего делать при выборе, кроме заполнения поля
    )

    btn_delete = ctk.CTkButton(win, text="Удалить", command=delete_film, fg_color="red", hover_color="darkred")
    btn_delete.place(x=100, y=250)  # или ниже, например y=270  # отступ снизу, чтобы кнопка не перекрывалась выпадающим списком

    win.mainloop()

def import_from_excel():
    try:
        # Указываем, что первая строка - это заголовки (header=0)
        df = pd.read_excel("Список пленок и толщин.xlsx", sheet_name="Перечень пленок", header=0)
        
        conn = sqlite3.connect("films.db")
        cur = conn.cursor()
        
        imported_count = 0
        updated_count = 0
        skipped_count = 0
        
        for index, row in df.iterrows():
            # Добавлена проверка на наличие значения 'ДА' в 16-й колонке (индекс 15)
            # Приводим к строке и верхнему регистру, чтобы избежать ошибок
            status = str(row.iloc[15]).strip().upper()
            if status != "ДА":
                skipped_count += 1
                continue
                
            sidak_raw = row.iloc[0]
            supplier_raw = row.iloc[3]
            thickness_raw = row.iloc[10]

            if pd.isna(sidak_raw) or pd.isna(supplier_raw) or pd.isna(thickness_raw):
                skipped_count += 1
                continue
            
            sidak = str(sidak_raw).strip()
            supplier = str(supplier_raw).strip()
            thickness_str = str(thickness_raw).strip().replace(",", ".")
            
            try:
                thickness = float(thickness_str)
            except ValueError:
                skipped_count += 1
                continue

            cur.execute("SELECT * FROM films WHERE sidak_num = ?", (sidak,))
            exists = cur.fetchone()

            if exists:
                cur.execute("UPDATE films SET supplier_name = ?, thickness = ? WHERE sidak_num = ?", 
                            (supplier, thickness, sidak))
                updated_count += 1
            else:
                cur.execute("INSERT INTO films (sidak_num, supplier_name, thickness) VALUES (?, ?, ?)",
                            (sidak, supplier, thickness))
                imported_count += 1
        
        conn.commit()
        conn.close()
        messagebox.showinfo("Импорт завершен", 
                            f"Импортировано новых плёнок: {imported_count}\n"
                            f"Обновлено существующих плёнок: {updated_count}\n"
                            f"Пропущено некорректных или неактуальных строк: {skipped_count}")
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл 'Список пленок и толщин.xlsx' не найден.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при импорте: {e}")

def edit_film_window():
    def on_film_selected(sidak_num):
        if sidak_num:
            details = get_film_details(sidak_num)
            if details:
                entry_edit_supplier.delete(0, ctk.END)
                entry_edit_supplier.insert(0, details[0])
                entry_edit_thickness.delete(0, ctk.END)
                entry_edit_thickness.insert(0, details[1])
            else:
                entry_edit_supplier.delete(0, ctk.END)
                entry_edit_thickness.delete(0, ctk.END)

    def update_film():
        sidak = entry_sidak.get().strip()
        supplier = entry_edit_supplier.get().strip()
        try:
            thickness = float(entry_edit_thickness.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Толщина должна быть числом")
            return

        if not sidak:
            messagebox.showerror("Ошибка", "Выберите плёнку для редактирования")
            return

        conn = sqlite3.connect("films.db")
        cur = conn.cursor()
        try:
            # Проверяем, существует ли плёнка
            cur.execute("SELECT 1 FROM films WHERE sidak_num = ?", (sidak,))
            exists = cur.fetchone()

            if not exists:
                messagebox.showerror("Ошибка", f"Плёнка с номером {sidak} не найдена в базе данных")
                return

            # Обновляем, если существует
            cur.execute("UPDATE films SET supplier_name = ?, thickness = ? WHERE sidak_num = ?",
                        (supplier, thickness, sidak))
            conn.commit()
            messagebox.showinfo("Успех", f"Данные для плёнки {sidak} обновлены")
            win.destroy()
        except sqlite3.Error as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        finally:
            conn.close()
    def check_film_exists(event=None):
        sidak = entry_sidak.get().strip()
        if not sidak:
            btn_save.configure(state="disabled")
            return

        conn = sqlite3.connect("films.db")
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM films WHERE sidak_num = ?", (sidak,))
        exists = cur.fetchone()
        conn.close()

        if exists:
            btn_save.configure(state="normal")
        else:
            btn_save.configure(state="disabled")
    win = ctk.CTkToplevel()
    win.title("Редактировать плёнку")
    win.geometry("400x350")

    ctk.CTkLabel(win, text="Введите или выберите номер Sidak").place(x=50, y=20)

    def on_film_selected_and_check(sidak_num):
        on_film_selected(sidak_num)        # заполняем поля
        check_film_exists()                # проверяем существование и активируем кнопку

    entry_sidak, results_frame = create_filterable_sidak_selector(
        parent=win,
        initial_x=50,
        initial_y=50,
        width=300,
        on_select_callback=on_film_selected_and_check
    )

    ctk.CTkLabel(win, text="Поставщик").place(x=50, y=120)
    entry_edit_supplier = ctk.CTkEntry(win, width=300)
    entry_edit_supplier.place(x=50, y=150)

    ctk.CTkLabel(win, text="Толщина").place(x=50, y=190)
    entry_edit_thickness = ctk.CTkEntry(win, width=300)
    entry_edit_thickness.place(x=50, y=220)

    btn_save = ctk.CTkButton(win, text="Сохранить изменения", command=update_film, fg_color="blue", hover_color="darkblue", state="disabled")
    btn_save.place(x=100, y=270)

    # Привязываем проверку к изменению поля ввода
    entry_sidak.bind("<KeyRelease>", check_film_exists)
    # Проверяем при открытии окна
    check_film_exists()

    win.mainloop()

def create_filterable_sidak_selector(parent, initial_x, initial_y, width, on_select_callback):
    """
    Создаёт фильтруемый выпадающий список Sidak номеров, как в liquid.py
    :param parent: родительское окно
    :param initial_x: x-координата поля ввода (относительно parent)
    :param initial_y: y-координата поля ввода
    :param width: ширина поля ввода и выпадающего списка
    :param on_select_callback: функция, вызываемая при выборе элемента (принимает sidak_num)
    :return: (entry_widget, results_frame) — поле ввода и фрейм для результатов
    """
    # Поле ввода
    entry_sidak = ctk.CTkEntry(parent, width=width)
    entry_sidak.place(x=initial_x, y=initial_y)

    # Фрейм для выпадающего списка
    results_frame = ctk.CTkScrollableFrame(parent, width=width, height=200)

    # Глобальные флаги
    is_dropdown_open = False

    def select_from_list(sidak_num):
        nonlocal is_dropdown_open
        entry_sidak.delete(0, 'end')
        entry_sidak.insert(0, sidak_num)
        results_frame.place_forget()
        is_dropdown_open = False
        if on_select_callback:
            on_select_callback(sidak_num)

    def filter_sidak(event=None):
        nonlocal is_dropdown_open
        text = entry_sidak.get().strip().lower()

        # Очищаем выпадающий список
        for widget in results_frame.winfo_children():
            widget.destroy()

        # Фильтруем данные из базы
        try:
            conn = sqlite3.connect("films.db")
            cur = conn.cursor()
            cur.execute("SELECT sidak_num FROM films WHERE lower(sidak_num) LIKE ? ORDER BY sidak_num", (f'%{text}%',))
            filtered_sidak = [r[0] for r in cur.fetchall()]
            conn.close()
        except sqlite3.Error as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к базе данных: {e}")
            return

        if not filtered_sidak:
            results_frame.place_forget()
            is_dropdown_open = False
            return

        # Получаем абсолютные координаты поля ввода ОТНОСИТЕЛЬНО ОКНА
        x_pos = initial_x
        y_pos = initial_y + entry_sidak.winfo_height() + 5  # +5 для отступа

        # Размещаем выпадающий список
        results_frame.place(x=x_pos, y=y_pos)
        results_frame.lift()  # ← ПОДНИМАЕМ НА ВЕРХНИЙ СЛОЙ!
        is_dropdown_open = True

        for sidak in filtered_sidak:
            btn = ctk.CTkButton(results_frame, text=sidak, command=lambda s=sidak: select_from_list(s))
            btn.pack(fill="x", pady=2)

    # Привязываем событие ввода
    entry_sidak.bind("<KeyRelease>", filter_sidak)

    return entry_sidak, results_frame