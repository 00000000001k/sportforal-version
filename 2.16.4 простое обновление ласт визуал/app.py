# app.py

import customtkinter as ctk
import tkinter.messagebox as messagebox

from globals import version, name_prog, document_blocks
from state_manager import (
    save_current_state,
    setup_auto_save,
    load_application_state,
    restore_application_state
)
from generation import combined_generation_process, get_all_placeholders_from_blocks
from events import add_event, remove_tab
from document_block import create_document_fields_block
import koshtorys
from excel_export import export_document_data_to_excel
from generation import generate_documents_word
from template_loader import get_available_templates

# === ДОДАНО ДЛЯ АВТОМАТИЧНИХ ОНОВЛЕНЬ ===
from ctk_update_manager import setup_auto_updates


# === КІНЕЦЬ ДОДАВАННЯ ===


def get_current_dynamic_fields(tabview):
    """Отримує динамічні поля для поточного заходу"""
    current_event = tabview.get()
    if not current_event:
        return []

    relevant_blocks = [b for b in document_blocks if b.get("tab_name") == current_event]
    if not relevant_blocks:
        return []

    dynamic_fields = get_all_placeholders_from_blocks(relevant_blocks)
    return dynamic_fields


def launch_main_app():
    global main_app_root, tabview, event_name_entry, update_manager

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("green")

    main_app_root = ctk.CTk()
    main_app_root.title(name_prog + version)
    main_app_root.geometry("1200x750")

    update_manager = setup_auto_updates(main_app_root, version)

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

    ctk.CTkButton(event_input_frame, text="➕ Додати", width=80, command=on_add_event_from_entry).pack(side="left", padx=5)

    # Інші кнопки (без шаблонів і додавання договору)
    ctk.CTkButton(top_controls_frame, text="Згенерувати договори",
                  command=lambda: generate_documents_word(tabview)).pack(padx=5, pady=5, side="left")

    ctk.CTkButton(top_controls_frame, text="💰 Кошторис",
                  command=lambda: koshtorys.fill_koshtorys(document_blocks),
                  fg_color="#FF9800").pack(side="left", padx=5)

    def export_excel_with_dynamic_fields():
        try:
            dynamic_fields = get_current_dynamic_fields(tabview)
            if not dynamic_fields:
                messagebox.showwarning("Увага", "Не знайдено полів для експорту. Додайте спочатку договори з шаблонами.")
                return
            success = export_document_data_to_excel(document_blocks, dynamic_fields)
            if success:
                messagebox.showinfo("Успіх", "Excel файл створено успішно!")
            else:
                messagebox.showerror("Помилка", "Не вдалося створити Excel файл")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка при експорті Excel: {e}")

    ctk.CTkButton(top_controls_frame, text="📥 Excel",
                  command=export_excel_with_dynamic_fields,
                  fg_color="#00BCD4").pack(side="left", padx=5)

    def check_updates():
        try:
            update_manager.check_updates_manual()
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося перевірити оновлення: {e}")

    ctk.CTkButton(top_controls_frame, text="🔄 Оновлення",
                  command=check_updates,
                  fg_color="#9C27B0", width=100).pack(side="left", padx=5)

    def show_dynamic_fields():
        try:
            dynamic_fields = get_current_dynamic_fields(tabview)
            if dynamic_fields:
                fields_text = "\n".join([f"• {field}" for field in dynamic_fields])
                messagebox.showinfo("Знайдені поля", f"Динамічні поля в поточному заході:\n\n{fields_text}")
            else:
                messagebox.showinfo("Поля", "Не знайдено жодних полів. Додайте договори з шаблонами.")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка при отриманні полів: {e}")

    ctk.CTkButton(top_controls_frame, text="🔍 Поля",
                  command=show_dynamic_fields,
                  fg_color="#607D8B", width=80).pack(side="left", padx=5)

    ctk.CTkLabel(top_controls_frame, text=version, text_color="gray", font=("Arial", 12)).pack(side="right", padx=10)

    setup_auto_save(main_app_root, document_blocks, tabview)

    return main_app_root, tabview


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