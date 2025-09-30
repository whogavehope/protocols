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
    cur.execute("SELECT supplier_name, thickness, heating_score, coffee_score, oil_score, hardness FROM films WHERE sidak_num = ?", (sidak,))
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
        
        # Получаем значения новых полей
        heating_score = combo_heating.get()
        heating_score = int(heating_score) if heating_score else None
        
        coffee_score = combo_coffee.get()
        coffee_score = int(coffee_score) if coffee_score else None
        
        oil_score = combo_oil.get()
        oil_score = int(oil_score) if oil_score else None
        
        hardness = combo_hardness.get()
        hardness = hardness if hardness != "" else None

        if not sidak or not supplier:
            messagebox.showerror("Ошибка", "Поля 'Номер Sidak' и 'Поставщик' обязательны")
            return

        conn = sqlite3.connect("films.db")
        cur = conn.cursor()
        try:
            cur.execute("""
                INSERT INTO films 
                (sidak_num, supplier_name, thickness, heating_score, coffee_score, oil_score, hardness) 
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (sidak, supplier, thickness, heating_score, coffee_score, oil_score, hardness))
            conn.commit()
            messagebox.showinfo("Успех", f"Плёнка {sidak} добавлена")
            win.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Ошибка", "Такая плёнка уже существует")
        finally:
            conn.close()

    win = ctk.CTkToplevel()
    win.title("Добавить плёнку")
    win.geometry("400x500")

    # Список вариантов для полей
    score_options = ["", "1", "2", "3", "4", "5"]
    hardness_options = ["", "5B", "4B", "3B", "2B", "B", "HB", "F", "H", "2H", "3H", "4H", "5H"]

    row = 0
    ctk.CTkLabel(win, text="Номер Sidak").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    entry_sidak = ctk.CTkEntry(win, width=300)
    entry_sidak.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkLabel(win, text="Поставщик").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    entry_supplier = ctk.CTkEntry(win, width=300)
    entry_supplier.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkLabel(win, text="Толщина").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    entry_thickness = ctk.CTkEntry(win, width=300)
    entry_thickness.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkLabel(win, text="Нагрев баллы").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    combo_heating = ctk.CTkComboBox(win, values=score_options, width=300)
    combo_heating.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkLabel(win, text="Кофе баллы").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    combo_coffee = ctk.CTkComboBox(win, values=score_options, width=300)
    combo_coffee.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkLabel(win, text="Масло баллы").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    combo_oil = ctk.CTkComboBox(win, values=score_options, width=300)
    combo_oil.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkLabel(win, text="Твердость").grid(row=row, column=0, sticky="w", padx=10, pady=5)
    combo_hardness = ctk.CTkComboBox(win, values=hardness_options, width=300)
    combo_hardness.grid(row=row, column=1, padx=5, pady=5)
    
    row += 1
    ctk.CTkButton(win, text="Сохранить", command=save_film, fg_color="green", hover_color="darkgreen").grid(row=row, columnspan=2, pady=20)
    
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
            
            # Новые поля из соответствующих столбцов
            heating_score_raw = row.iloc[11]  # 12-й столбец (индекс 11)
            coffee_score_raw = row.iloc[12]   # 13-й столбец (индекс 12)
            oil_score_raw = row.iloc[13]      # 14-й столбец (индекс 13)
            hardness_raw = row.iloc[9]        # 10-й столбец (индекс 9) - hardness

            if pd.isna(sidak_raw) or pd.isna(supplier_raw) or pd.isna(thickness_raw):
                skipped_count += 1
                continue
            
            sidak = str(sidak_raw).strip()
            supplier = str(supplier_raw).strip()
            thickness_str = str(thickness_raw).strip().replace(",", ".")
            
            # Обработка новых полей
            heating_score = None
            if not pd.isna(heating_score_raw):
                try:
                    heating_score = int(float(heating_score_raw))
                except (ValueError, TypeError):
                    heating_score = None
            
            coffee_score = None
            if not pd.isna(coffee_score_raw):
                try:
                    coffee_score = int(float(coffee_score_raw))
                except (ValueError, TypeError):
                    coffee_score = None
            
            oil_score = None
            if not pd.isna(oil_score_raw):
                try:
                    oil_score = int(float(oil_score_raw))
                except (ValueError, TypeError):
                    oil_score = None
            
            hardness = None
            if not pd.isna(hardness_raw):
                hardness_str = str(hardness_raw).strip().upper()
                # Проверяем, что значение находится в допустимом списке
                valid_hardness = ['5B', '4B', '3B', '2B', 'B', 'HB', 'F', 'H', '2H', '3H', '4H', '5H']
                if hardness_str in valid_hardness:
                    hardness = hardness_str

            try:
                thickness = float(thickness_str)
            except ValueError:
                skipped_count += 1
                continue

            cur.execute("SELECT * FROM films WHERE sidak_num = ?", (sidak,))
            exists = cur.fetchone()

            if exists:
                cur.execute("""
                    UPDATE films SET 
                    supplier_name = ?, 
                    thickness = ?,
                    heating_score = ?,
                    coffee_score = ?,
                    oil_score = ?,
                    hardness = ?
                    WHERE sidak_num = ?
                """, (supplier, thickness, heating_score, coffee_score, oil_score, hardness, sidak))
                updated_count += 1
            else:
                cur.execute("""
                    INSERT INTO films 
                    (sidak_num, supplier_name, thickness, heating_score, coffee_score, oil_score, hardness) 
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (sidak, supplier, thickness, heating_score, coffee_score, oil_score, hardness))
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
                
                # Заполняем новые поля
                combo_edit_heating.set(str(details[2]) if details[2] is not None else "")
                combo_edit_coffee.set(str(details[3]) if details[3] is not None else "")
                combo_edit_oil.set(str(details[4]) if details[4] is not None else "")
                combo_edit_hardness.set(details[5] if details[5] is not None else "")
            else:
                entry_edit_supplier.delete(0, ctk.END)
                entry_edit_thickness.delete(0, ctk.END)
                combo_edit_heating.set("")
                combo_edit_coffee.set("")
                combo_edit_oil.set("")
                combo_edit_hardness.set("")

    def update_film():
        sidak = entry_sidak.get().strip()
        supplier = entry_edit_supplier.get().strip()
        try:
            thickness = float(entry_edit_thickness.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Толщина должна быть числом")
            return
        
        # Получаем значения новых полей
        heating_score = combo_edit_heating.get()
        heating_score = int(heating_score) if heating_score else None
        
        coffee_score = combo_edit_coffee.get()
        coffee_score = int(coffee_score) if coffee_score else None
        
        oil_score = combo_edit_oil.get()
        oil_score = int(oil_score) if oil_score else None
        
        hardness = combo_edit_hardness.get()
        hardness = hardness if hardness != "" else None

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

            # Обновляем все поля
            cur.execute("""
                UPDATE films SET supplier_name = ?, thickness = ?, heating_score = ?, 
                coffee_score = ?, oil_score = ?, hardness = ? 
                WHERE sidak_num = ?
            """, (supplier, thickness, heating_score, coffee_score, oil_score, hardness, sidak))
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
    win.geometry("400x550")

    # Список вариантов для полей
    score_options = ["", "1", "2", "3", "4", "5"]
    hardness_options = ["", "5B", "4B", "3B", "2B", "B", "HB", "F", "H", "2H", "3H", "4H", "5H"]

    ctk.CTkLabel(win, text="Введите или выберите номер Sidak").place(x=50, y=20)

    def on_film_selected_and_check(sidak_num):
        on_film_selected(sidak_num)
        check_film_exists()

    entry_sidak, results_frame = create_filterable_sidak_selector(
        parent=win,
        initial_x=50,
        initial_y=50,
        width=300,
        on_select_callback=on_film_selected_and_check
    )

    y_pos = 120
    ctk.CTkLabel(win, text="Поставщик").place(x=50, y=y_pos)
    entry_edit_supplier = ctk.CTkEntry(win, width=300)
    entry_edit_supplier.place(x=50, y=y_pos + 30)

    y_pos += 70
    ctk.CTkLabel(win, text="Толщина").place(x=50, y=y_pos)
    entry_edit_thickness = ctk.CTkEntry(win, width=300)
    entry_edit_thickness.place(x=50, y=y_pos + 30)

    y_pos += 70
    ctk.CTkLabel(win, text="Нагрев баллы").place(x=50, y=y_pos)
    combo_edit_heating = ctk.CTkComboBox(win, values=score_options, width=300)
    combo_edit_heating.place(x=50, y=y_pos + 30)

    y_pos += 70
    ctk.CTkLabel(win, text="Кофе баллы").place(x=50, y=y_pos)
    combo_edit_coffee = ctk.CTkComboBox(win, values=score_options, width=300)
    combo_edit_coffee.place(x=50, y=y_pos + 30)

    y_pos += 70
    ctk.CTkLabel(win, text="Масло баллы").place(x=50, y=y_pos)
    combo_edit_oil = ctk.CTkComboBox(win, values=score_options, width=300)
    combo_edit_oil.place(x=50, y=y_pos + 30)

    y_pos += 70
    ctk.CTkLabel(win, text="Твердость").place(x=50, y=y_pos)
    combo_edit_hardness = ctk.CTkComboBox(win, values=hardness_options, width=300)
    combo_edit_hardness.place(x=50, y=y_pos + 30)

    btn_save = ctk.CTkButton(win, text="Сохранить изменения", command=update_film, 
                           fg_color="blue", hover_color="darkblue", state="disabled")
    btn_save.place(x=100, y=y_pos + 70)

    entry_sidak.bind("<KeyRelease>", check_film_exists)
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



