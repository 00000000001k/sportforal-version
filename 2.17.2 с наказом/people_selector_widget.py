# people_selector_widget.py

import customtkinter as ctk
import tkinter.messagebox as messagebox
from people_manager import people_manager, PEOPLE_CONFIG, SPECIAL_ROLES


class PeopleSelectorDialog:
    def __init__(self, parent):
        self.parent = parent
        self.dialog = None
        self.checkboxes = {}
        self.special_role_vars = {}

    def show_dialog(self):
        """Показує діалог вибору людей"""
        self.dialog = ctk.CTkToplevel(self.parent)
        self.dialog.title("Вибір осіб для документів")
        self.dialog.geometry("500x600")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()

        # Центруємо діалог
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (600 // 2)
        self.dialog.geometry(f"500x600+{x}+{y}")

        self._create_widgets()

    def _create_widgets(self):
        """Створює віджети діалогу"""
        # Заголовок
        title_label = ctk.CTkLabel(
            self.dialog,
            text="Оберіть осіб для включення в документи",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=10)

        # Основна область з прокруткою
        main_frame = ctk.CTkScrollableFrame(self.dialog, height=400)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Розділ 1: Основні особи
        self._create_main_people_section(main_frame)

        # Розділювач
        separator = ctk.CTkFrame(main_frame, height=2)
        separator.pack(fill="x", pady=15)

        # Розділ 2: Спеціальні ролі
        self._create_special_roles_section(main_frame)

        # Кнопки
        self._create_buttons()

    def _create_main_people_section(self, parent):
        """Створює розділ основних осіб"""
        section_label = ctk.CTkLabel(
            parent,
            text="Основні особи:",
            font=("Arial", 14, "bold"),
            anchor="w"
        )
        section_label.pack(fill="x", pady=(0, 10))

        for person_id, person_data in PEOPLE_CONFIG.items():
            # Фрейм для кожної особи
            person_frame = ctk.CTkFrame(parent, fg_color="transparent")
            person_frame.pack(fill="x", pady=5)

            # Чекбокс
            var = ctk.BooleanVar()
            var.set(people_manager.is_person_selected(person_id))

            checkbox = ctk.CTkCheckBox(
                person_frame,
                text=person_data['display_name'],
                variable=var,
                command=lambda pid=person_id: self._on_person_toggle(pid)
            )
            checkbox.pack(anchor="w", padx=10, pady=5)

            self.checkboxes[person_id] = var

            # Інформація про посаду (якщо потрібно)
            if len(person_data['position']) > 50:  # Довга посада
                position_label = ctk.CTkLabel(
                    person_frame,
                    text=f"Посада: {person_data['position'][:50]}...",
                    font=("Arial", 10),
                    text_color="gray",
                    anchor="w"
                )
                position_label.pack(anchor="w", padx=30)

    def _create_special_roles_section(self, parent):
        """Створює розділ спеціальних ролей"""
        section_label = ctk.CTkLabel(
            parent,
            text="Спеціальні ролі:",
            font=("Arial", 14, "bold"),
            anchor="w"
        )
        section_label.pack(fill="x", pady=(0, 10))

        for role_id, role_config in SPECIAL_ROLES.items():
            # Фрейм для ролі
            role_frame = ctk.CTkFrame(parent)
            role_frame.pack(fill="x", pady=10, padx=5)

            # Заголовок ролі
            role_title = ctk.CTkLabel(
                role_frame,
                text=role_config['title'],
                font=("Arial", 12, "bold")
            )
            role_title.pack(pady=(10, 5))

            # Радіо кнопки для вибору
            var = ctk.StringVar()
            current_selection = people_manager.get_special_role(role_id)
            var.set(current_selection)

            for option_person_id in role_config['options']:
                person_data = PEOPLE_CONFIG.get(option_person_id)
                if person_data:
                    radio = ctk.CTkRadioButton(
                        role_frame,
                        text=person_data['display_name'],
                        variable=var,
                        value=option_person_id,
                        command=lambda rid=role_id, pid=option_person_id: self._on_special_role_change(rid, pid)
                    )
                    radio.pack(anchor="w", padx=20, pady=2)

            self.special_role_vars[role_id] = var

    def _create_buttons(self):
        """Створює кнопки діалогу"""
        button_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
        button_frame.pack(fill="x", padx=20, pady=10)

        # Кнопка "Скасувати"
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="Скасувати",
            command=self._on_cancel,
            fg_color="gray"
        )
        cancel_btn.pack(side="right", padx=5)

        # Кнопка "Зберегти"
        save_btn = ctk.CTkButton(
            button_frame,
            text="Зберегти",
            command=self._on_save,
            fg_color="#2E8B57"
        )
        save_btn.pack(side="right", padx=5)

        # Кнопка "Попередній перегляд"
        preview_btn = ctk.CTkButton(
            button_frame,
            text="Попередній перегляд",
            command=self._on_preview,
            fg_color="#FF9800"
        )
        preview_btn.pack(side="left", padx=5)

        # Кнопка "Очистити все"
        clear_btn = ctk.CTkButton(
            button_frame,
            text="Очистити все",
            command=self._on_clear_all,
            fg_color="#DC3545"
        )
        clear_btn.pack(side="left", padx=5)

    def _on_person_toggle(self, person_id):
        """Обробник перемикання чекбоксу особи"""
        people_manager.toggle_person(person_id)

    def _on_special_role_change(self, role_id, person_id):
        """Обробник зміни спеціальної ролі"""
        people_manager.set_special_role(role_id, person_id)

    def _on_preview(self):
        """Показує попередній перегляд вибраних осіб"""
        summary = people_manager.get_summary()
        replacements = people_manager.generate_replacements()

        preview_text = f"{summary}\n\n"
        preview_text += "Згенеровані плейсхолдери:\n"

        for placeholder, replacement in replacements.items():
            preview_text += f"\n{placeholder}:\n{replacement}\n{'-' * 50}\n"

        if not replacements:
            preview_text += "Жодних плейсхолдерів не згенеровано (не обрано жодної особи)"

        # Створюємо вікно попереднього перегляду
        preview_dialog = ctk.CTkToplevel(self.dialog)
        preview_dialog.title("Попередній перегляд")
        preview_dialog.geometry("600x500")
        preview_dialog.transient(self.dialog)

        # Текстове поле з прокруткою
        text_widget = ctk.CTkTextbox(preview_dialog)
        text_widget.pack(fill="both", expand=True, padx=20, pady=20)
        text_widget.insert("1.0", preview_text)
        text_widget.configure(state="disabled")

        # Кнопка закриття
        close_btn = ctk.CTkButton(
            preview_dialog,
            text="Закрити",
            command=preview_dialog.destroy
        )
        close_btn.pack(pady=10)

    def _on_clear_all(self):
        """Очищує всі вибори"""
        result = messagebox.askyesno(
            "Підтвердження",
            "Очистити всі вибори осіб?"
        )

        if result:
            # Очищуємо основних осіб
            people_manager.selected_people.clear()

            # Скидаємо спеціальні ролі до значень за замовчуванням
            for role_id, role_config in SPECIAL_ROLES.items():
                people_manager.set_special_role(role_id, role_config['default'])

            # Оновлюємо інтерфейс
            for person_id, var in self.checkboxes.items():
                var.set(False)

            for role_id, var in self.special_role_vars.items():
                var.set(SPECIAL_ROLES[role_id]['default'])

            people_manager.save_settings()

    def _on_save(self):
        """Зберігає налаштування та закриває діалог"""
        people_manager.save_settings()
        messagebox.showinfo("Успіх", "Налаштування осіб збережено!")
        self.dialog.destroy()

    def _on_cancel(self):
        """Скасовує зміни та закриває діалог"""
        # Відновлюємо попередні налаштування
        people_manager.load_settings()
        self.dialog.destroy()


class PeopleSelectorButton:
    """Кнопка для відкриття діалогу вибору людей"""

    def __init__(self, parent):
        self.parent = parent
        self.button = ctk.CTkButton(
            parent,
            text="👥 Налаштування осіб",
            command=self.show_people_selector,
            fg_color="#9C27B0",
            hover_color="#7B1FA2",
            width=150
        )
        self.update_button_text()

    def show_people_selector(self):
        """Відкриває діалог вибору людей"""
        dialog = PeopleSelectorDialog(self.parent)
        dialog.show_dialog()

        # Оновлюємо текст кнопки після закриття діалогу
        self.parent.after(100, self.update_button_text)

    def update_button_text(self):
        """Оновлює текст кнопки з кількістю обраних осіб"""
        count = people_manager.get_selected_count()
        if count > 0:
            self.button.configure(text=f"👥 Особи ({count})")
        else:
            self.button.configure(text="👥 Налаштування осіб")

    def pack(self, **kwargs):
        """Пакує кнопку"""
        self.button.pack(**kwargs)

    def grid(self, **kwargs):
        """Розміщує кнопку в сітці"""
        self.button.grid(**kwargs)