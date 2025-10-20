# app.py

import customtkinter as ctk
import tkinter.messagebox as messagebox

from globals import version,name_prog, document_blocks, FIELDS
from state_manager import (
    save_current_state,
    setup_auto_save,
    load_application_state,
    restore_application_state
)
from generation import combined_generation_process
from events import add_event, remove_tab
from document_block import create_document_fields_block
import koshtorys
from excel_export import export_document_data_to_excel
from generation import generate_documents_word


def launch_main_app():
    global main_app_root, tabview, event_name_entry

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("green")

    main_app_root = ctk.CTk()
    main_app_root.title(name_prog + version)
    main_app_root.geometry("1200x750")

    # --- Верхня панель ---
    top_controls_frame = ctk.CTkFrame(main_app_root)
    top_controls_frame.pack(pady=10, padx=10, fill="x")

    # Tabview (вкладки заходів)
    tabview = ctk.CTkTabview(main_app_root)
    tabview.pack(fill="both", expand=True, padx=10, pady=10)

    # Завантаження та відновлення попереднього стану
    saved_state = load_application_state()
    if saved_state:
        restore_application_state(saved_state, tabview, main_app_root)
    else:
        # Якщо немає збереженого стану, створюємо базову структуру
        print("[INFO] Запуск з чистого листа")

    # Введення назви заходу
    event_input_frame = ctk.CTkFrame(top_controls_frame)
    event_input_frame.pack(side="left", fill="x", expand=True, padx=5)

    event_name_entry = ctk.CTkEntry(event_input_frame, placeholder_text="Назва заходу", width=250)
    event_name_entry.pack(side="left", padx=5, fill="x", expand=True)

    def on_add_event_from_entry():
        name = event_name_entry.get().strip()
        if name:
            add_event(name, tabview)
            event_name_entry.delete(0, 'end')
            save_current_state(document_blocks, tabview)
        else:
            messagebox.showwarning("Увага", "Назва заходу не може бути порожньою!")

    ctk.CTkButton(event_input_frame, text="➕ Додати", width=80, command=on_add_event_from_entry) \
        .pack(side="left", padx=5)

    # Кнопки дій
    # ctk.CTkButton(top_controls_frame, text="➕ Додати договір",
    #               command=lambda: add_contract_to_current_event(tabview),
    #               fg_color="#2196F3").pack(side="left", padx=5)

    ctk.CTkButton(top_controls_frame, text="Згенерувати договори",
                  command=lambda: generate_documents_word(tabview)).pack(padx=5, pady=5, side="left")

    ctk.CTkButton(top_controls_frame, text="💰 Кошторис",
                  command=lambda: koshtorys.fill_koshtorys(document_blocks),
                  fg_color="#FF9800").pack(side="left", padx=5)

    ctk.CTkButton(top_controls_frame, text="📥 Excel",
                  command=lambda: export_document_data_to_excel(document_blocks, FIELDS),
                  fg_color="#00BCD4").pack(side="left", padx=5)


    ctk.CTkLabel(top_controls_frame, text=version, text_color="gray", font=("Arial", 12)) \
        .pack(side="right", padx=10)

    # Налаштування автоматичного збереження при закритті програми
    setup_auto_save(main_app_root, document_blocks, tabview)

    return main_app_root, tabview


def add_contract_to_current_event(tabview):
    """Додає новий договір до поточного заходу"""
    selected_tab_name = tabview.get()
    if not selected_tab_name:
        messagebox.showwarning("Увага", "Спочатку створіть захід!")
        return

    tab = tabview.tab(selected_tab_name)
    if hasattr(tab, "contracts_frame"):
        create_document_fields_block(tab.contracts_frame, tabview)
        # Автоматично зберігаємо стан після додавання договору
        save_current_state(document_blocks, tabview)
        print("[INFO] Договір додано та стан збережено")
    else:
        messagebox.showerror("Помилка", "Не знайдено контейнер для договорів")


def auto_save_on_change():
    """Функція для автоматичного збереження при змінах"""
    try:
        save_current_state(document_blocks, tabview)
        print("[DEBUG] Автозбереження виконано")
    except Exception as e:
        print(f"[ERROR] Помилка автозбереження: {e}")


# Додаткові функції для керування станом програми
def manual_save():
    """Ручне збереження стану"""
    try:
        save_current_state(document_blocks, tabview)
        messagebox.showinfo("Успіх", "Стан програми збережено успішно!")
    except Exception as e:
        messagebox.showerror("Помилка", f"Не вдалося зберегти стан:\n{e}")


def clear_all_data():
    """Очищення всіх даних (з підтвердженням)"""
    result = messagebox.askyesno(
        "Підтвердження",
        "Ви впевнені, що хочете очистити всі дані?\nЦю дію неможливо скасувати!"
    )

    if result:
        from state_manager import clear_saved_state
        clear_saved_state()
        messagebox.showinfo("Інформація", "Збережені дані очищено. Перезапустіть програму.")


if __name__ == "__main__":
    # Запуск головної програми
    app_root, app_tabview = launch_main_app()

    # Збереження глобальних посилань для автозбереження
    main_app_root = app_root
    tabview = app_tabview

    # Запуск головного циклу
    app_root.mainloop()