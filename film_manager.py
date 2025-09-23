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
        selected_film = combo_films.get()
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

    ctk.CTkLabel(win, text="Выберите плёнку").pack(pady=5)
    
    films = get_all_films()
    combo_films = ctk.CTkComboBox(win, values=films, width=300)
    combo_films.pack(pady=5)

    btn_delete = ctk.CTkButton(win, text="Удалить", command=delete_film, fg_color="red", hover_color="darkred")
    btn_delete.pack(pady=10)

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
    def on_film_selected(choice):
        selected_sidak = choice
        if selected_sidak:
            details = get_film_details(selected_sidak)
            if details:
                entry_edit_supplier.delete(0, ctk.END)
                entry_edit_supplier.insert(0, details[0])
                entry_edit_thickness.delete(0, ctk.END)
                entry_edit_thickness.insert(0, details[1])
            else:
                entry_edit_supplier.delete(0, ctk.END)
                entry_edit_thickness.delete(0, ctk.END)
    
    def update_film():
        sidak = combo_edit_films.get().strip()
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
            cur.execute("UPDATE films SET supplier_name = ?, thickness = ? WHERE sidak_num = ?",
                        (supplier, thickness, sidak))
            conn.commit()
            messagebox.showinfo("Успех", f"Данные для плёнки {sidak} обновлены")
            win.destroy()
        except sqlite3.Error as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        finally:
            conn.close()

    win = ctk.CTkToplevel()
    win.title("Редактировать плёнку")

    ctk.CTkLabel(win, text="Выберите плёнку").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    
    films = get_all_films()
    combo_edit_films = ctk.CTkComboBox(win, values=films, width=300, command=on_film_selected)
    combo_edit_films.grid(row=0, column=1, padx=5, pady=5)

    ctk.CTkLabel(win, text="Поставщик").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    entry_edit_supplier = ctk.CTkEntry(win, width=300)
    entry_edit_supplier.grid(row=1, column=1, padx=5, pady=5)

    ctk.CTkLabel(win, text="Толщина").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    entry_edit_thickness = ctk.CTkEntry(win, width=300)
    entry_edit_thickness.grid(row=2, column=1, padx=5, pady=5)

    ctk.CTkButton(win, text="Сохранить изменения", command=update_film, fg_color="blue", hover_color="darkblue").grid(row=3, columnspan=2, pady=10)