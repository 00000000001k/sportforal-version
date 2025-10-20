# events.py - виправлена версія

import customtkinter as ctk
import tkinter.messagebox as messagebox
from template_loader import get_available_templates
from document_block import create_document_fields_block
from globals import document_blocks
from state_manager import save_current_state


def get_event_number_dialog(parent):
    """Диалог для ввода номера заходу"""
    dialog = ctk.CTkToplevel(parent)
    dialog.title("Номер заходу")
    dialog.geometry("300x200")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()

    # Центрируем окно
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
    y = (dialog.winfo_screenheight() // 2) - (200 // 2)
    dialog.geometry(f"300x200+{x}+{y}")

    result = None

    # Основной фрейм
    main_frame = ctk.CTkFrame(dialog)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Метка
    ctk.CTkLabel(main_frame, text="Номер заходу:", font=("Arial", 12)).pack(pady=10)

    # Поле ввода
    entry = ctk.CTkEntry(main_frame, width=200, placeholder_text="Наприклад: 14")
    entry.pack(pady=10)
    entry.focus()

    # Кнопки
    button_frame = ctk.CTkFrame(main_frame)
    button_frame.pack(pady=10)

    def on_ok():
        nonlocal result
        text = entry.get().strip()
        if text:
            try:
                result = int(text)
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Помилка", "Введіть коректний номер!")
        else:
            messagebox.showerror("Помилка", "Номер не може бути пустим!")

    def on_cancel():
        nonlocal result
        result = None
        dialog.destroy()

    # Кнопки Ок та Скасувати
    ctk.CTkButton(
        button_frame,
        text="OK",
        command=on_ok,
        width=80
    ).pack(side="left", padx=5)

    ctk.CTkButton(
        button_frame,
        text="Скасувати",
        command=on_cancel,
        width=100,
        fg_color="#a6a6a6",  # светло-серый
        hover_color="#8c8c8c",  # чуть темнее при наведении
        text_color="black"
    ).pack(side="left", padx=10)

    # Обработка Enter
    def on_enter(event):
        on_ok()

    entry.bind("<Return>", on_enter)

    dialog.wait_window()
    return result


def add_event(event_name, tabview, restore=False, event_number=None):
    """Додає новий захід з панеллю вибору шаблонів

    Args:
        event_name: назва заходу
        tabview: об'єкт TabView
        restore: чи це відновлення стану (за замовчуванням False)
        event_number: номер заходу (за замовчуванням None)
    """

    if event_name in [tabview.tab(tab_name) for tab_name in tabview._tab_dict.keys()]:
        if not restore:  # Показуємо попередження тільки якщо це не відновлення
            messagebox.showwarning("Увага", f"Захід '{event_name}' вже існує!")
            return

    # Якщо це не відновлення і номер не заданий, запитуємо у користувача
    if not restore and event_number is None:
        event_number = get_event_number_dialog(tabview.master)
        if event_number is None:  # Користувач скасував
            return

    # Створюємо нову вкладку
    tab = tabview.add(event_name)

    # === БЛОК ЗАГАЛЬНИХ ДАНИХ ЗАХОДУ ===
    from event_common_fields import create_common_fields_block
    common_frame, common_entries = create_common_fields_block(tab, event_name, tabview)

    # === Панель керування шаблонами для цього заходу ===
    template_control_frame = ctk.CTkFrame(tab)
    template_control_frame.pack(fill="x", padx=10, pady=(10, 5))

    # Мітка
    ctk.CTkLabel(template_control_frame, text="Оберіть шаблон:", font=("Arial", 12, "bold")).pack(side="left",
                                                                                                  padx=(10, 5))

    # Випадаючий список шаблонів з автооновленням
    template_var = ctk.StringVar()

    # Спочатку ініціалізуємо шаблони
    initial_templates = get_available_templates()

    # Створюємо меню
    template_menu = ctk.CTkOptionMenu(
        template_control_frame,
        variable=template_var,
        values=list(initial_templates.keys()) if initial_templates else ["Немає шаблонів"],
        width=200
    )
    template_menu.pack(side="left", padx=5)

    # Встановлюємо початкове значення
    if initial_templates:
        template_var.set(list(initial_templates.keys())[0])
    else:
        template_var.set("Немає шаблонів")

    # Тепер визначаємо функцію оновлення (після створення menu)
    def refresh_templates():
        """Оновлює список шаблонів"""
        try:
            templates_dict = get_available_templates()
            if templates_dict:
                template_names = list(templates_dict.keys())
                template_menu.configure(values=template_names)
                if not template_var.get() or template_var.get() not in template_names:
                    template_var.set(template_names[0])
                return templates_dict
            else:
                template_menu.configure(values=["Немає шаблонів"])
                template_var.set("Немає шаблонів")
                return {}
        except Exception as e:
            print(f"[ERROR] Помилка оновлення шаблонів: {e}")
            template_menu.configure(values=["Помилка завантаження"])
            template_var.set("Помилка завантаження")
            return {}

    # Зберігаємо початкові шаблони
    templates_dict = initial_templates

    # Кнопка оновлення шаблонів
    def on_refresh_templates():
        nonlocal templates_dict
        templates_dict = refresh_templates()
        if not restore:  # Показуємо повідомлення тільки якщо це не відновлення
            messagebox.showinfo("Оновлено", f"Знайдено {len(templates_dict)} шаблонів")

    # Кнопка оновлення (без tooltip_text)
    refresh_button = ctk.CTkButton(
        template_control_frame,
        text="🔄",
        width=30,
        height=30,
        command=on_refresh_templates
    )
    refresh_button.pack(side="left", padx=2)

    # Кнопка додавання договору
    def add_contract_to_this_event():
        """Додає договір до поточного заходу"""
        selected_template = template_var.get()

        if not selected_template or selected_template in ["Немає шаблонів", "Помилка завантаження"]:
            messagebox.showwarning("Увага", "Спочатку оберіть валідний шаблон!")
            return

        # Оновлюємо шаблони на всякий випадок
        current_templates = get_available_templates()
        template_path = current_templates.get(selected_template)

        if not template_path:
            messagebox.showerror("Помилка", f"Шаблон '{selected_template}' не знайдено! Спробуйте оновити список.")
            return

        # Додаємо договір
        create_document_fields_block(contracts_frame, tabview, template_path)
        print(f"[INFO] Договір додано з шаблоном: {selected_template}")

    ctk.CTkButton(
        template_control_frame,
        text="➕ Додати шаблон",
        command=add_contract_to_this_event,
        fg_color="#2E8B57",
        hover_color="#228B22"
    ).pack(side="left", padx=10)

    # Інформаційна мітка з номером заходу
    info_text = f"Шаблонів знайдено: {len(templates_dict)}"
    if event_number is not None:
        info_text += f"  |  Номер заходу: {event_number}"

    info_label = ctk.CTkLabel(
        template_control_frame,
        text=info_text,
        text_color="gray60",
        font=("Arial", 10)
    )
    info_label.pack(side="left", padx=10)

    # === Контейнер для договорів ===
    contracts_frame = ctk.CTkScrollableFrame(tab)
    contracts_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Зберігаємо посилання на фрейм для подальшого використання
    tab.contracts_frame = contracts_frame
    tab.common_frame = common_frame
    tab.common_entries = common_entries
    tab.template_var = template_var
    tab.templates_dict = templates_dict
    tab.refresh_templates = refresh_templates
    tab.event_number = event_number  # Зберігаємо номер заходу

    if not restore:  # Логуємо тільки якщо це не відновлення
        print(
            f"[INFO] Створено захід '{event_name}' з номером {event_number} та {len(templates_dict)} доступними шаблонами")


def remove_tab(tab_name, tabview):
    """Видаляє вкладку заходу"""
    result = messagebox.askyesno(
        "Підтвердження",
        f"Ви впевнені, що хочете видалити захід '{tab_name}'?\n"
        "Всі договори цього заходу будуть втрачені!"
    )

    if result:
        # Видаляємо всі блоки документів цього заходу
        global document_blocks
        document_blocks = [block for block in document_blocks if block.get("tab_name") != tab_name]

        # Видаляємо вкладку
        tabview.delete(tab_name)

        # Зберігаємо стан
        save_current_state(document_blocks, tabview)

        print(f"[INFO] Захід '{tab_name}' видалено")
        messagebox.showinfo("Успіх", f"Захід '{tab_name}' видалено успішно!")


def get_event_contracts_count(event_name):
    """Повертає кількість договорів у заході"""
    return len([block for block in document_blocks if block.get("tab_name") == event_name])


def get_all_events():
    """Повертає список всіх заходів"""
    # Це потрібно реалізувати відповідно до структури tabview
    pass