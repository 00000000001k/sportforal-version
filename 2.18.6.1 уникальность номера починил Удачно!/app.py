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
from generation import get_all_placeholders_from_blocks
from events import add_event, remove_tab
from document_block import create_document_fields_block
import koshtorys
from excel_export import export_document_data_to_excel
from generation import generate_documents_word
from template_loader import get_available_templates
from people_selector_widget import PeopleSelectorButton

# === ДОДАНО ДЛЯ АВТОМАТИЧНИХ ОНОВЛЕНЬ ===
from ctk_update_manager import setup_auto_updates



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

    # === ПАНЕЛЬ КЕРУВАННЯ ЗАХОДАМИ ===
    event_control_frame = ctk.CTkFrame(top_controls_frame)
    event_control_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))

    # Підпанель для додавання заходу
    event_input_frame = ctk.CTkFrame(event_control_frame, fg_color="transparent")
    event_input_frame.pack(fill="x", pady=(5, 0))

    ctk.CTkLabel(event_input_frame, text="Керування заходами:",
                 font=("Arial", 12, "bold")).pack(anchor="w", padx=5)

    add_event_frame = ctk.CTkFrame(event_input_frame, fg_color="transparent")
    add_event_frame.pack(fill="x", pady=5)

    event_name_entry = ctk.CTkEntry(add_event_frame, placeholder_text="Назва нового заходу", width=200)
    event_name_entry.pack(side="left", padx=5, fill="x", expand=True)

    def on_add_event_from_entry():
        name = event_name_entry.get().strip()
        if name:
            add_event(name, tabview)
            event_name_entry.delete(0, 'end')
            save_current_state(document_blocks, tabview)
        else:
            messagebox.showwarning("Увага", "Назва заходу не може бути порожньою!")

    ctk.CTkButton(add_event_frame, text="➕ Додати захід", width=120,
                  command=on_add_event_from_entry,
                  fg_color="#2E8B57", hover_color="#228B22").pack(side="left", padx=5)

    # Підпанель для видалення заходу
    delete_event_frame = ctk.CTkFrame(event_input_frame, fg_color="transparent")
    delete_event_frame.pack(fill="x", pady=(5, 5))

    def on_remove_current_event():
        """Видаляє поточний активний захід"""
        current_tab = tabview.get()
        if not current_tab:
            messagebox.showwarning("Увага", "Немає активного заходу для видалення!")
            return

        # Перевіряємо, чи є договори в цьому заході
        contracts_count = len([block for block in document_blocks if block.get("tab_name") == current_tab])

        warning_text = f"Ви впевнені, що хочете видалити захід '{current_tab}'?"
        if contracts_count > 0:
            warning_text += f"\n\nУ цьому заході є {contracts_count} договорів, які також будуть видалені!"
        warning_text += "\n\nЦю дію неможливо скасувати!"

        result = messagebox.askyesno("Підтвердження видалення", warning_text)

        if result:
            remove_tab(current_tab, tabview)

    def get_all_events():
        """Отримує список всіх заходів"""
        try:
            return list(tabview._tab_dict.keys()) if hasattr(tabview, '_tab_dict') else []
        except:
            return []

    def on_remove_selected_event():
        """Видаляє обраний захід зі списку"""
        all_events = get_all_events()

        if not all_events:
            messagebox.showinfo("Інформація", "Немає заходів для видалення!")
            return

        # Створюємо діалогове вікно для вибору заходу
        dialog = ctk.CTkToplevel(main_app_root)
        dialog.title("Видалити захід")
        dialog.geometry("400x300")
        dialog.transient(main_app_root)
        dialog.grab_set()

        # Центруємо діалог
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")

        ctk.CTkLabel(dialog, text="Оберіть захід для видалення:",
                     font=("Arial", 14, "bold")).pack(pady=10)

        # Список заходів з інформацією про кількість договорів
        events_frame = ctk.CTkScrollableFrame(dialog, height=150)
        events_frame.pack(fill="both", expand=True, padx=20, pady=10)

        selected_event = ctk.StringVar()

        for event_name in all_events:
            contracts_count = len([block for block in document_blocks if block.get("tab_name") == event_name])
            event_text = f"{event_name}"
            if contracts_count > 0:
                event_text += f" ({contracts_count} договорів)"

            event_radio = ctk.CTkRadioButton(events_frame, text=event_text,
                                             variable=selected_event, value=event_name)
            event_radio.pack(anchor="w", pady=2, padx=10)

        # Кнопки
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=20, pady=10)

        def confirm_delete():
            event_to_delete = selected_event.get()
            if not event_to_delete:
                messagebox.showwarning("Увага", "Оберіть захід для видалення!")
                return

            contracts_count = len([block for block in document_blocks if block.get("tab_name") == event_to_delete])

            warning_text = f"Видалити захід '{event_to_delete}'?"
            if contracts_count > 0:
                warning_text += f"\n\nРазом з {contracts_count} договорами!"

            if messagebox.askyesno("Остаточне підтвердження", warning_text):
                dialog.destroy()
                remove_tab(event_to_delete, tabview)

        def cancel_delete():
            dialog.destroy()

        ctk.CTkButton(buttons_frame, text="❌ Видалити",
                      command=confirm_delete,
                      fg_color="#DC3545", hover_color="#C82333").pack(side="left", padx=5)

        ctk.CTkButton(buttons_frame, text="Скасувати",
                      command=cancel_delete).pack(side="right", padx=5)

    # Кнопки видалення
    ctk.CTkButton(delete_event_frame, text="🗑 Видалити захiд", width=140,
                  command=on_remove_current_event,
                  fg_color="#DC3545", hover_color="#C82333").pack(side="left", padx=5)


    # === ПАНЕЛЬ ОСНОВНИХ ФУНКЦІЙ ===
    main_functions_frame = ctk.CTkFrame(top_controls_frame)
    main_functions_frame.pack(side="right", padx=(10, 0))

    # Перший ряд кнопок
    first_row = ctk.CTkFrame(main_functions_frame, fg_color="transparent")
    first_row.pack(pady=(5, 5))

    ctk.CTkButton(first_row, text="📄 Згенерувати договори",
                  command=lambda: generate_documents_word(tabview),
                  width=180).pack(side="left", padx=5)

    ctk.CTkButton(first_row, text="💰 Кошторис",
                  command=lambda: koshtorys.fill_koshtorys(document_blocks),
                  fg_color="#FF9800", width=120).pack(side="left", padx=5)

    # Другий ряд кнопок
    second_row = ctk.CTkFrame(main_functions_frame, fg_color="transparent")
    second_row.pack(pady=5)

    def export_excel_with_dynamic_fields():
        try:
            dynamic_fields = get_current_dynamic_fields(tabview)
            if not dynamic_fields:
                messagebox.showwarning("Увага",
                                       "Не знайдено полів для експорту. Додайте спочатку договори з шаблонами.")
                return
            success = export_document_data_to_excel(document_blocks, dynamic_fields)
            if success:
                messagebox.showinfo("Успіх", "Excel файл створено успішно!")
            else:
                messagebox.showerror("Помилка", "Не вдалося створити Excel файл")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка при експорті Excel: {e}")

    ctk.CTkButton(second_row, text="📥 Excel",
                  command=export_excel_with_dynamic_fields,
                  fg_color="#00BCD4", width=80).pack(side="left", padx=5)

    def check_updates():
        try:
            update_manager.check_updates_manual()
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося перевірити оновлення: {e}")

    ctk.CTkButton(second_row, text="🔄 Оновлення",
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

    ctk.CTkButton(second_row, text="🔍 Поля",
                  command=show_dynamic_fields,
                  fg_color="#607D8B", width=80).pack(side="left", padx=5)

    # Третій ряд - версія
    third_row = ctk.CTkFrame(main_functions_frame, fg_color="transparent")
    third_row.pack(pady=5)

    people_button = PeopleSelectorButton(third_row)
    people_button.pack(side="left", padx=5)

    # Переносимо версію в четвертий ряд
    version_frame = ctk.CTkFrame(main_functions_frame, fg_color="transparent")
    version_frame.pack(pady=(5, 5))

    ctk.CTkLabel(version_frame, text=version, text_color="gray", font=("Arial", 12)).pack()


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