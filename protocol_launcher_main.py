import customtkinter as ctk
from tkinter import messagebox, ttk
from nagrev_gpt import show_input_form as show_nagrev_protocol
from liquid import show_input_form as show_liquid_protocol
from film_manager import add_film_window, delete_film_window, import_from_excel, edit_film_window

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

# Настройка внешнего вида CustomTkinter
ctk.set_appearance_mode("light")  # Режим: "light" (светлый), "dark" (тёмный), "system" (системный)
ctk.set_default_color_theme("blue")  # Темы: "blue" (по умолчанию), "green", "dark-blue"

root = ctk.CTk()
root.title("Выбор протокола")
root.geometry("400x350")

lbl = ctk.CTkLabel(root, text="Выберите вид протокола", font=("Arial", 12))
lbl.pack(pady=10)

combo = ctk.CTkComboBox(root, values=list(PROTOCOLS.keys()), width=40)
combo.set("Протокол по локальному нагреву") # Устанавливаем значение по умолчанию
combo.pack(pady=5)

btn_launch = ctk.CTkButton(
    root,
    text="Запустить",
    command=launch_selected_protocol,
    width=20,
    fg_color="green",
    hover_color="darkgreen"
)
btn_launch.pack(pady=10)

btn_add = ctk.CTkButton(
    root,
    text="Добавить плёнку",
    command=launch_add_film,
    width=20,
    fg_color="blue",
    hover_color="darkblue"
)
btn_add.pack(pady=5)

btn_delete = ctk.CTkButton(
    root,
    text="Удалить плёнку",
    command=launch_delete_film,
    width=20,
    fg_color="red",
    hover_color="darkred"
)
btn_delete.pack(pady=5)

btn_import = ctk.CTkButton(
    root,
    text="Импортировать из Excel",
    command=launch_import_from_excel,
    width=25,
    fg_color="gray",
    hover_color="darkgray"
)
btn_import.pack(pady=5)

btn_edit = ctk.CTkButton(
    root,
    text="Редактировать плёнку",
    command=launch_edit_film,
    width=25,
    fg_color="yellow",
    hover_color="darkyellow",
    text_color="black"
)
btn_edit.pack(pady=5)

root.mainloop()