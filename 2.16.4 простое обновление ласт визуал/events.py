# events.py

import customtkinter as ctk
import tkinter.messagebox as messagebox
from template_loader import get_available_templates
from document_block import create_document_fields_block
from globals import document_blocks
from state_manager import save_current_state


def add_event(event_name, tabview):
    """Додає новий захід з панеллю вибору шаблонів"""

    if event_name in [tabview.tab(tab_name) for tab_name in tabview._tab_dict.keys()]:
        messagebox.showwarning("Увага", f"Захід '{event_name}' вже існує!")
        return

    # Створюємо нову вкладку
    tab = tabview.add(event_name)

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
        messagebox.showinfo("Оновлено", f"Знайдено {len(templates_dict)} шаблонів")

    # ИСПРАВЛЕНО: убрал tooltip_text
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
        text="➕ Додати договір",
        command=add_contract_to_this_event,
        fg_color="#2E8B57",
        hover_color="#228B22"
    ).pack(side="left", padx=10)

    # Інформаційна мітка
    info_label = ctk.CTkLabel(
        template_control_frame,
        text=f"Шаблонів знайдено: {len(templates_dict)}",
        text_color="gray60",
        font=("Arial", 10)
    )
    info_label.pack(side="left", padx=10)

    # === Контейнер для договорів ===
    contracts_frame = ctk.CTkScrollableFrame(tab, label_text=f"Договори заходу: {event_name}")
    contracts_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Зберігаємо посилання на фрейм для подальшого використання
    tab.contracts_frame = contracts_frame
    tab.template_var = template_var
    tab.templates_dict = templates_dict
    tab.refresh_templates = refresh_templates

    print(f"[INFO] Створено захід '{event_name}' з {len(templates_dict)} доступними шаблонами")


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