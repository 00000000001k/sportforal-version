# events.py - версія з унікальними заходами
import os

import customtkinter as ctk
import tkinter.messagebox as messagebox
from template_loader import get_available_templates
from document_block import create_document_fields_block
from globals import document_blocks
from state_manager import save_current_state, get_existing_events


def get_event_dialog(parent):
    """Розширений діалог для створення/відновлення заходу"""
    dialog = ctk.CTkToplevel(parent)
    dialog.title("Керування заходами")
    dialog.geometry("500x600")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()

    # Центрируем окно
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
    y = (dialog.winfo_screenheight() // 2) - (600 // 2)
    dialog.geometry(f"500x600+{x}+{y}")

    result = {"action": None, "event_number": None, "event_name": None}

    # Получаем существующие события
    existing_events = get_existing_events()

    # Основной фрейм
    main_frame = ctk.CTkScrollableFrame(dialog)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # === СЕКЦІЯ ІСНУЮЧИХ ЗАХОДІВ ===
    existing_frame = ctk.CTkFrame(main_frame)
    existing_frame.pack(fill="x", pady=(0, 20))

    ctk.CTkLabel(existing_frame, text="Існуючі заходи:", font=("Arial", 14, "bold")).pack(pady=(10, 5))

    if existing_events:
        # Создаем таблицу существующих событий
        events_text = ctk.CTkTextbox(existing_frame, height=150, font=("Consolas", 11))
        events_text.pack(fill="x", padx=10, pady=(0, 10))

        # Заполняем таблицу с улучшенным заголовком
        header = f"{'№':^5} | {'Назва заходу':^25} | {'Договорів':^8} | {'Статус':^12}\n"
        separator = "-" * 65 + "\n"
        events_text.insert("0.0", header + separator)

        for event_num, event_data in sorted(existing_events.items()):
            event_name = event_data.get('name', 'Без назви')
            contracts_count = len(event_data.get('contracts', []))
            has_tab = event_data.get('has_tab', True)

            # Определяем статус
            if has_tab:
                status = "Відкрито"
                status_color = ""
            else:
                status = "Тільки дані"
                status_color = " ⚠️"

            # Обрезаем длинные названия
            display_name = event_name[:25] if len(event_name) <= 25 else event_name[:22] + "..."

            line = f"{event_num:^5} | {display_name:^25} | {contracts_count:^8} | {status:^12}{status_color}\n"
            events_text.insert("end", line)



            # Додаємо пояснення статусів
            explanation_frame = ctk.CTkFrame(existing_frame)
            explanation_frame.pack(fill="x", padx=10, pady=(5, 10))

            explanation_text = (
                "Статуси: 'Відкрито' - захід має активну вкладку; "
                "'Тільки дані' ⚠️ - збережені дані без вкладки"
            )

            ctk.CTkLabel(
                explanation_frame,
                text=explanation_text,
                font=("Arial", 9),
                text_color="gray60",
                wraplength=450
            ).pack(pady=5)




        events_text.configure(state="disabled")

    else:
        no_events_label = ctk.CTkLabel(
            existing_frame,
            text="Немає збережених заходів",
            text_color="gray60",
            font=("Arial", 11, "italic")
        )
        no_events_label.pack(pady=10)

    # === СЕКЦІЯ СТВОРЕННЯ НОВОГО ЗАХОДУ ===
    new_event_frame = ctk.CTkFrame(main_frame)
    new_event_frame.pack(fill="x", pady=(0, 20))

    ctk.CTkLabel(new_event_frame, text="Створити новий захід:", font=("Arial", 14, "bold")).pack(pady=(10, 5))

    # Поле номера
    number_frame = ctk.CTkFrame(new_event_frame)
    number_frame.pack(fill="x", padx=10, pady=5)

    ctk.CTkLabel(number_frame, text="Номер заходу:", font=("Arial", 12)).pack(side="left", padx=(10, 5))
    number_entry = ctk.CTkEntry(number_frame, width=100, placeholder_text="Наприклад: 14")
    number_entry.pack(side="left", padx=5)

    # Автоподсказка по номеру
    number_hint_label = ctk.CTkLabel(number_frame, text="", text_color="orange", font=("Arial", 10))
    number_hint_label.pack(side="left", padx=10)

    # Поле названия
    name_frame = ctk.CTkFrame(new_event_frame)
    name_frame.pack(fill="x", padx=10, pady=5)

    ctk.CTkLabel(name_frame, text="Назва заходу:", font=("Arial", 12)).pack(side="left", padx=(10, 5))
    name_entry = ctk.CTkEntry(name_frame, width=300, placeholder_text="Наприклад: День спорту")
    name_entry.pack(side="left", padx=5)

    # Функция проверки номера при вводе
    def check_number_hint(*args):
        try:
            number_text = number_entry.get().strip()
            if number_text and number_text.isdigit():
                event_num = int(number_text)
                if event_num in existing_events:
                    existing_name = existing_events[event_num].get('name', 'Без назви')
                    number_hint_label.configure(text=f"⚠️ Вже існує: {existing_name}")
                    number_hint_label.configure(text_color="red")
                else:
                    number_hint_label.configure(text="✓ Номер вільний")
                    number_hint_label.configure(text_color="green")
            else:
                number_hint_label.configure(text="")
        except:
            number_hint_label.configure(text="")

    # Привязываем проверку к изменению текста
    number_entry.bind('<KeyRelease>', check_number_hint)

    # Кнопки для нового события
    new_buttons_frame = ctk.CTkFrame(new_event_frame)
    new_buttons_frame.pack(pady=10)

    def create_new_event():
        number_text = number_entry.get().strip()
        name_text = name_entry.get().strip()

        if not number_text:
            messagebox.showerror("Помилка", "Введіть номер заходу!")
            return

        if not name_text:
            messagebox.showerror("Помилка", "Введіть назву заходу!")
            return

        try:
            event_number = int(number_text)
        except ValueError:
            messagebox.showerror("Помилка", "Номер заходу має бути числом!")
            return

        # Проверяем уникальность номера
        if event_number in existing_events:
            existing_name = existing_events[event_number].get('name', 'Без назви')
            messagebox.showerror(
                "Номер вже використовується",
                f"Захід з номером {event_number} вже існує:\n'{existing_name}'\n\nОберіть інший номер."
            )
            return

        result["action"] = "create"
        result["event_number"] = event_number
        result["event_name"] = name_text
        dialog.destroy()

    ctk.CTkButton(
        new_buttons_frame,
        text="✓ Створити захід",
        command=create_new_event,
        fg_color="#2E8B57",
        hover_color="#228B22"
    ).pack(side="left", padx=5)

    # === СЕКЦІЯ ВІДНОВЛЕННЯ ЗАХОДУ ===
    if existing_events:
        restore_frame = ctk.CTkFrame(main_frame)
        restore_frame.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(restore_frame, text="Відновити існуючий захід:", font=("Arial", 14, "bold")).pack(pady=(10, 5))

        # Поле для номера відновлення
        restore_number_frame = ctk.CTkFrame(restore_frame)
        restore_number_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(restore_number_frame, text="Номер для відновлення:", font=("Arial", 12)).pack(side="left",
                                                                                                   padx=(10, 5))
        restore_entry = ctk.CTkEntry(restore_number_frame, width=100, placeholder_text="Номер")
        restore_entry.pack(side="left", padx=5)

        # Подсказка по восстановлению
        restore_hint_label = ctk.CTkLabel(restore_number_frame, text="", text_color="blue", font=("Arial", 10))
        restore_hint_label.pack(side="left", padx=10)

        def check_restore_hint(*args):
            try:
                restore_text = restore_entry.get().strip()
                if restore_text and restore_text.isdigit():
                    event_num = int(restore_text)
                    if event_num in existing_events:
                        existing_name = existing_events[event_num].get('name', 'Без назви')
                        restore_hint_label.configure(text=f"↻ Знайдено: {existing_name}")
                        restore_hint_label.configure(text_color="blue")
                    else:
                        restore_hint_label.configure(text="❌ Не знайдено")
                        restore_hint_label.configure(text_color="red")
                else:
                    restore_hint_label.configure(text="")
            except:
                restore_hint_label.configure(text="")

        restore_entry.bind('<KeyRelease>', check_restore_hint)

        def restore_event():
            restore_text = restore_entry.get().strip()

            if not restore_text:
                messagebox.showerror("Помилка", "Введіть номер заходу для відновлення!")
                return

            try:
                event_number = int(restore_text)
            except ValueError:
                messagebox.showerror("Помилка", "Номер заходу має бути числом!")
                return

            if event_number not in existing_events:
                messagebox.showerror("Не знайдено", f"Захід з номером {event_number} не знайдено у збережених даних!")
                return

            event_data = existing_events[event_number]
            result["action"] = "restore"
            result["event_number"] = event_number
            result["event_name"] = event_data.get('name', f'Захід {event_number}')
            dialog.destroy()

        restore_buttons_frame = ctk.CTkFrame(restore_frame)
        restore_buttons_frame.pack(pady=10)

        ctk.CTkButton(
            restore_buttons_frame,
            text="↻ Відновити захід",
            command=restore_event,
            fg_color="#1f538d",
            hover_color="#14396d"
        ).pack(side="left", padx=5)

    # === КНОПКА СКАСУВАННЯ ===
    cancel_frame = ctk.CTkFrame(main_frame)
    cancel_frame.pack(fill="x", pady=(10, 0))

    def cancel_dialog():
        result["action"] = None
        dialog.destroy()

    ctk.CTkButton(
        cancel_frame,
        text="Скасувати",
        command=cancel_dialog,
        fg_color="#a6a6a6",
        hover_color="#8c8c8c",
        text_color="black"
    ).pack(pady=10)

    # Обработка Enter для создания
    def on_create_enter(event):
        create_new_event()

    # Обработка Enter для восстановления
    def on_restore_enter(event):
        if existing_events:
            restore_event()

    name_entry.bind("<Return>", on_create_enter)
    number_entry.bind("<Return>", on_create_enter)
    if existing_events:
        restore_entry.bind("<Return>", on_restore_enter)

    # Фокус на первое поле
    number_entry.focus()

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

    # Проверяем существование вкладки
    if event_name in [tabview.tab(tab_name) for tab_name in tabview._tab_dict.keys()]:
        if not restore:  # Показуємо попередження тільки якщо це не відновлення
            messagebox.showwarning("Увага", f"Захід '{event_name}' вже існує!")
            return

    # Якщо це не відновлення і номер не заданий, викликаємо новий діалог
    if not restore and event_number is None:
        dialog_result = get_event_dialog(tabview.master)

        if dialog_result["action"] == "create":
            event_number = dialog_result["event_number"]
            event_name = dialog_result["event_name"]
        elif dialog_result["action"] == "restore":
            # Восстанавливаем событие через state_manager
            from state_manager import restore_single_event
            success = restore_single_event(dialog_result["event_number"], tabview)
            if success:
                messagebox.showinfo("Успіх", f"Захід №{dialog_result['event_number']} відновлено успішно!")
            else:
                messagebox.showerror("Помилка", "Не вдалося відновити захід!")
            return
        else:
            return  # Пользователь отменил

    # Створюємо нову вкладку
    tab = tabview.add(event_name)

    # === БЛОК ЗАГАЛЬНИХ ДАНИХ ЗАХОДУ ===
    from event_common_fields import create_common_fields_block
    # ВАЖЛИВО: Передаємо event_number при створенні загальних полів
    common_frame, common_entries = create_common_fields_block(tab, event_name, tabview, event_number=event_number)

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

        # Видаляємо захід зі збереженого стану


        print(f"[INFO] Захід '{tab_name}' видалено")
        messagebox.showinfo("Успіх", f"Захід '{tab_name}' видалено успішно!")


def get_event_contracts_count(event_name):
    """Повертає кількість договорів у заході"""
    return len([block for block in document_blocks if block.get("tab_name") == event_name])


def get_all_events():
    """Повертає список всіх заходів"""
    # Це потрібно реалізувати відповідно до структури tabview
    pass