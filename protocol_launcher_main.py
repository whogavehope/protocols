import customtkinter as ctk
from tkinter import messagebox
from nagrev_gpt import show_input_form as show_nagrev_protocol
from liquid import show_input_form as show_liquid_protocol
from film_manager import add_film_window, delete_film_window, import_from_excel, edit_film_window

import sys
import os

# =============== ПАТЧ ДЛЯ ПЕРЕХВАТА AFTER ===============
_all_after_ids = set()
_original_after = None
_original_after_cancel = None

def patched_after(self, ms, func=None, *args):
    """Перехватываем after, сохраняем ID для отмены"""
    if func is None:
        after_id = _original_after(self, ms)
    else:
        after_id = _original_after(self, ms, func, *args)
    _all_after_ids.add(after_id)
    return after_id

def patched_after_cancel(self, after_id):
    """Перехватываем after_cancel, удаляем ID из списка"""
    if after_id in _all_after_ids:
        _all_after_ids.remove(after_id)
    return _original_after_cancel(self, after_id)

def cancel_all_remaining_after():
    """Отменяем все after-задачи перед закрытием"""
    for after_id in list(_all_after_ids):
        try:
            root.after_cancel(after_id)
        except:
            pass
    _all_after_ids.clear()
# =======================================================

PROTOCOLS = {
    "Протокол по локальному нагреву": show_nagrev_protocol,
    "Протокол испытания на воздействие жидкостей": show_liquid_protocol,
}

def launch_selected_protocol():
    selected = combo.get()
    if selected not in PROTOCOLS:
        messagebox.showwarning("Внимание", "Выберите протокол")
        return
    root.destroy()
    PROTOCOLS[selected]()

def launch_add_film():
    add_film_window()

def launch_delete_film():
    delete_film_window()

def launch_import_from_excel():
    import_from_excel()

def launch_edit_film():
    edit_film_window()

# Настройка внешнего вида
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Создание главного окна
root = ctk.CTk()
root.title("Выбор протокола")
root.geometry("400x350")

# === 🔥 ГЛАВНОЕ ИСПРАВЛЕНИЕ: НЕ ПЕРЕДАЁМ ROOT ВРУЧНУЮ! ===
_original_after = root.after
_original_after_cancel = root.after_cancel

# ✅ Правильно: self приходит из root.after(...)
root.after = lambda *args: patched_after(*args)
root.after_cancel = lambda after_id: patched_after_cancel(*args)  # ← тоже без root!

# =========================================================

# Элементы интерфейса
lbl = ctk.CTkLabel(root, text="Выберите вид протокола", font=("Arial", 12))
lbl.pack(pady=10)

combo = ctk.CTkComboBox(root, values=list(PROTOCOLS.keys()), width=280)
combo.set("Протокол по локальному нагреву")
combo.pack(pady=5)

btn_launch = ctk.CTkButton(
    root,
    text="Запустить",
    command=launch_selected_protocol,
    width=200,
    fg_color="green",
    hover_color="darkgreen"
)
btn_launch.pack(pady=10)

btn_add = ctk.CTkButton(
    root,
    text="Добавить плёнку",
    command=launch_add_film,
    width=200,
    fg_color="blue",
    hover_color="darkblue"
)
btn_add.pack(pady=5)

btn_delete = ctk.CTkButton(
    root,
    text="Удалить плёнку",
    command=launch_delete_film,
    width=200,
    fg_color="red",
    hover_color="darkred"
)
btn_delete.pack(pady=5)

btn_import = ctk.CTkButton(
    root,
    text="Импортировать из Excel",
    command=launch_import_from_excel,
    width=200,
    fg_color="gray",
    hover_color="darkgray"
)
btn_import.pack(pady=5)

btn_edit = ctk.CTkButton(
    root,
    text="Редактировать плёнку",
    command=launch_edit_film,
    width=200,
    fg_color="yellow",
    hover_color="yellow",
    text_color="black"
)
btn_edit.pack(pady=5)

# Обработчик закрытия
def on_closing():
    cancel_all_remaining_after()
    root.quit()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# Запуск цикла
try:
    root.mainloop()
except Exception as e:
    error_str = str(e)
    if ("invalid command name" in error_str and
        any(keyword in error_str for keyword in ["after", "update", "click", "dpi", "animation"])):
        pass
    else:
        import traceback
        traceback.print_exc()