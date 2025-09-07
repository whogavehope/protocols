import tkinter as tk
from tkinter import ttk, messagebox

# Здесь подключаем функцию из существующего протокола
# Например, nagrev_gpt.py содержит show_input_form
from nagrev_gpt import show_input_form as show_nagrev_protocol

# Для остальных протоколов можно будет создать аналогичные функции:
# from liquid_test import show_input_form as show_liquid_protocol
# from hardness_test import show_input_form as show_hardness_protocol

PROTOCOLS = {
    "Протокол по локальному нагреву": show_nagrev_protocol,
    # "Протокол испытания на воздействие жидкостей": show_liquid_protocol,
    # "Протокол испытания на твердость покрытия": show_hardness_protocol,
}

def launch_selected_protocol():
    selected = combo.get()
    if selected not in PROTOCOLS:
        messagebox.showwarning("Внимание", "Выберите протокол")
        return
    root.destroy()  # закрываем окно выбора
    PROTOCOLS[selected]()  # запускаем выбранный протокол

root = tk.Tk()
root.title("Выбор протокола")
root.geometry("400x150")

lbl = tk.Label(root, text="Выберите вид протокола", font=("Arial", 12))
lbl.pack(pady=20)

combo = ttk.Combobox(root, values=list(PROTOCOLS.keys()), width=40)
combo.pack(pady=5)
combo.current(0)

btn = tk.Button(root, text="Запустить", command=launch_selected_protocol, width=20, bg="lightgreen")
btn.pack(pady=15)

root.mainloop()
