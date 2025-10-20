# sportforall/gui/dialogs.py

import tkinter as tk
from tkinter import ttk, messagebox

# Імпортуємо власні моделі, якщо потрібні (напр., для типізації)
# from sportforall.models import Event, Contract

def ask_event_name(master) -> str | None:
    """
    Відкриває діалог для введення назви нового заходу.

    Args:
        master: Батьківське вікно Tkinter.

    Returns:
        str: Назва заходу, якщо введено, або None, якщо скасовано.
    """
    # Перевіряємо, чи існує батьківське вікно перед створенням Toplevel
    if not master:
         print("Помилка: Немає головного вікна (master) для створення діалогу.")
         return None

    dialog = tk.Toplevel(master)
    dialog.title("Новий Захід")
    dialog.transient(master) # Робить діалог модальним відносно головного вікна
    dialog.grab_set() # Захоплює всі події введення
    dialog.resizable(False, False) # Забороняє зміну розміру

    # Розміщуємо діалог по центру батьківського вікна (опціонально)
    # dialog_width = 300
    # dialog_height = 120
    # parent_pos_x = master.winfo_x()
    # parent_pos_y = master.winfo_y()
    # parent_width = master.winfo_width()
    # parent_height = master.winfo_height()
    # dialog.geometry(f'{dialog_width}x{dialog_height}+{parent_pos_x + int(parent_width/2 - dialog_width/2)}+{parent_pos_y + int(parent_height/2 - dialog_height/2)}')


    label = ttk.Label(dialog, text="Введіть назву нового заходу:")
    label.pack(padx=10, pady=10)

    entry = ttk.Entry(dialog, width=40)
    entry.pack(padx=10, pady=5)
    entry.focus_set() # Встановлює фокус на поле введення при відкритті

    result = None # Змінна для зберігання результату

    def on_ok():
        nonlocal result
        event_name = entry.get().strip()
        if event_name:
            result = event_name
            dialog.destroy()
        else:
            messagebox.showwarning("Попередження", "Назва заходу не може бути порожньою.")

    def on_cancel():
        dialog.destroy()

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)
    ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Скасувати", command=on_cancel).pack(side="left", padx=5)

    # Прив'язуємо Enter до OK та Escape до Скасувати
    dialog.bind('<Return>', lambda e=None: on_ok())
    dialog.bind('<Escape>', lambda e=None: on_cancel())

    master.wait_window(dialog) # Чекаємо, поки діалог закриється
    return result


def ask_contract_name(master, event_name: str) -> str | None:
    """
    Відкриває діалог для введення назви нового договору.

    Args:
        master: Батьківське вікно Tkinter.
        event_name: Назва заходу, до якого додається договір (для відображення в заголовку).

    Returns:
        str: Назва договору, якщо введено, або None, якщо скасовано.
    """
    # Перевіряємо, чи існує батьківське вікно перед створенням Toplevel
    if not master:
         print("Помилка: Немає головного вікна (master) для створення діалогу.")
         return None

    dialog = tk.Toplevel(master)
    dialog.title(f"Новий Договір для '{event_name}'")
    dialog.transient(master) # Робить діалог модальним відносно головного вікна
    dialog.grab_set() # Захоплює всі події введення
    dialog.resizable(False, False) # Забороняє зміну розміру


    tk.Label(dialog, text="Введіть назву нового договору:").pack(padx=10, pady=5)
    entry = ttk.Entry(dialog, width=50)
    entry.pack(padx=10, pady=5)
    entry.focus_set() # Встановлює фокус на поле введення при відкритті

    result = None # Змінна для зберігання результату

    def on_ok():
        nonlocal result
        contract_name = entry.get().strip()
        if contract_name:
             result = contract_name
             dialog.destroy()
        else:
             messagebox.showwarning("Попередження", "Назва договору не може бути порожньою.")


    def on_cancel():
        dialog.destroy()

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)
    ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Скасувати", command=on_cancel).pack(side="left", padx=5)

    # Прив'язуємо Enter до OK та Escape до Скасувати
    dialog.bind('<Return>', lambda e=None: on_ok())
    dialog.bind('<Escape>', lambda e=None: on_cancel())

    master.wait_window(dialog) # Чекаємо, поки діалог буде закрито
    return result