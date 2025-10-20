# event_common_fields.py
# -*- coding: utf-8 -*-

import customtkinter as ctk
import tkinter.messagebox as messagebox

# Видаляємо проблемний імпорт - будемо імпортувати локально
# from app import add_contract_to_current_event
from globals import FIELDS, EXAMPLES
from custom_widgets import CustomEntry
from gui_utils import bind_entry_shortcuts, create_context_menu
from state_manager import save_current_state

# Поля, які є загальними для всього заходу
COMMON_FIELDS = ["захід", "дата", "адреса"]

# Глобальний словник для збереження загальних даних кожного заходу
event_common_data = {}


def create_common_fields_block(parent_frame, event_name, tabview=None):
    """Створює блок загальних полів для заходу"""

    # Основний фрейм для загальних полів
    common_frame = ctk.CTkFrame(parent_frame, border_width=2, border_color="green")
    common_frame.pack(pady=(5, 15), padx=5, fill="x")

    # Заголовок з кнопкою видалення
    header_frame = ctk.CTkFrame(common_frame, fg_color="transparent")
    header_frame.pack(fill="x", padx=10, pady=(10, 5))

    header_label = ctk.CTkLabel(header_frame,
                                text="📋 ЗАГАЛЬНІ ДАНІ ЗАХОДУ",
                                font=("Arial", 16, "bold"),
                                text_color="green")
    header_label.pack(side="top")

    # Кнопка видалення заходу
    def remove_current_event():
        if messagebox.askokcancel("Видалити захід",
                                  f"Ви дійсно бажаєте видалити захід '{event_name}'?\n\nБудуть видалені:\n• Всі договори цього заходу\n• Загальні дані заходу\n• Вкладка заходу"):
            from globals import document_blocks

            # Видаляємо всі блоки договорів цього заходу
            blocks_to_remove = [block for block in document_blocks if block.get("tab_name") == event_name]
            for block in blocks_to_remove:
                if block in document_blocks:
                    document_blocks.remove(block)
                # Знищуємо фрейм блоку
                if "frame" in block:
                    try:
                        block["frame"].destroy()
                    except:
                        pass

            # Видаляємо загальні дані заходу
            remove_event_common_data(event_name)

            # Видаляємо вкладку
            try:
                tabview.delete(event_name)
                print(f"[INFO] Захід '{event_name}' видалено успішно")
            except Exception as e:
                print(f"[ERROR] Помилка при видаленні вкладки '{event_name}': {e}")

            # Зберігаємо стан
            save_current_state(document_blocks, tabview)

    delete_button = ctk.CTkButton(header_frame,
                                  text="❌ Видалити захід",
                                  fg_color="red",
                                  hover_color="darkred",
                                  width=120,
                                  height=32,
                                  font=("Arial", 11, "bold"),
                                  command=remove_current_event)
    delete_button.pack(side="right")

    # Функція для додавання договору (локальний імпорт)
    def add_contract_handler():
        try:
            # Імпортуємо функцію локально, щоб уникнути циклічного імпорту
            from app import add_contract_to_current_event
            add_contract_to_current_event(tabview)
        except ImportError as e:
            print(f"[ERROR] Не вдалося імпортувати add_contract_to_current_event: {e}")
            messagebox.showerror("Помилка", "Не вдалося додати договір. Перевірте конфігурацію програми.")

    ctk.CTkButton(header_frame, text="➕ Додати договір",
                  command=add_contract_handler,
                  fg_color="#2196F3").pack(side="left", padx=5)

    # Підказка
    info_label = ctk.CTkLabel(common_frame,
                              text="Ці поля автоматично копіюються у всі договори цього заходу",
                              font=("Arial", 12),
                              text_color="gray60")
    info_label.pack(pady=(0, 10))

    # Фрейм для полів
    fields_frame = ctk.CTkFrame(common_frame, fg_color="transparent")
    fields_frame.pack(fill="x", padx=10, pady=(0, 10))

    # Створюємо поля
    common_entries = {}
    context_menu = create_context_menu(common_frame)

    for i, field_key in enumerate(COMMON_FIELDS):
        # Лейбл
        label = ctk.CTkLabel(fields_frame,
                             text=f"<{field_key}>",
                             anchor="w",
                             width=100,
                             font=("Arial", 12, "bold"))
        label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

        # Поле вводу
        entry = CustomEntry(fields_frame, field_name=field_key, examples_dict=EXAMPLES)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        fields_frame.columnconfigure(1, weight=1)

        # Кнопка підказки
        hint_button = ctk.CTkButton(fields_frame,
                                    text="ℹ",
                                    width=28,
                                    height=28,
                                    font=("Arial", 14),
                                    command=lambda h=EXAMPLES.get(field_key, "Немає підказки"), f=field_key:
                                    messagebox.showinfo(f"Підказка для <{f}>", h))
        hint_button.grid(row=i, column=2, padx=(0, 5), pady=5, sticky="e")

        # Зв'язуємо контекстне меню
        bind_entry_shortcuts(entry, context_menu)

        # Зберігаємо посилання на поле
        common_entries[field_key] = entry

        # Завантажуємо збережене значення
        if event_name in event_common_data and field_key in event_common_data[event_name]:
            entry.set_text(event_common_data[event_name][field_key])

        # Прив'язуємо функцію оновлення при зміні
        entry.bind("<KeyRelease>", lambda e, field=field_key: update_common_field(event_name, field, e))
        entry.bind("<FocusOut>", lambda e, field=field_key: update_common_field(event_name, field, e))

    # Ініціалізуємо дані для заходу, якщо їх немає
    if event_name not in event_common_data:
        event_common_data[event_name] = {}

    return common_frame, common_entries


def update_common_field(event_name, field_key, event):
    """Оновлює загальне поле та синхронізує його з усіма договорами заходу"""
    from globals import document_blocks

    # Отримуємо нове значення
    new_value = event.widget.get()

    # Зберігаємо у глобальних даних
    if event_name not in event_common_data:
        event_common_data[event_name] = {}
    event_common_data[event_name][field_key] = new_value

    # Синхронізуємо з усіма договорами цього заходу
    sync_common_fields_to_contracts(event_name)


def sync_common_fields_to_contracts(event_name):
    """Синхронізує загальні поля з усіма договорами заходу"""
    from globals import document_blocks

    if event_name not in event_common_data:
        return

    # Знаходимо всі блоки цього заходу
    event_blocks = [block for block in document_blocks if block.get("tab_name") == event_name]

    # Оновлюємо загільні поля в кожному блоці
    for block in event_blocks:
        entries = block.get("entries", {})
        for field_key in COMMON_FIELDS:
            if field_key in entries and field_key in event_common_data[event_name]:
                entry_widget = entries[field_key]
                # Зберігаємо поточний стан
                current_state = entry_widget.cget("state")
                # Тимчасovo робимо поле доступним для редагування
                entry_widget.configure(state="normal")
                # Оновлюємо значення
                entry_widget.set_text(event_common_data[event_name][field_key])
                # Повертаємо попередній стан
                entry_widget.configure(state=current_state)


def fill_common_fields_for_new_contract(event_name, contract_entries):
    """Заповнює загальні поля для нового договору"""
    if event_name not in event_common_data:
        return

    for field_key in COMMON_FIELDS:
        if field_key in contract_entries and field_key in event_common_data[event_name]:
            entry_widget = contract_entries[field_key]
            # Заповнюємо значенням із загальних даних
            entry_widget.set_text(event_common_data[event_name][field_key])


def get_common_data_for_event(event_name):
    """Повертає загальні дані для заходу"""
    return event_common_data.get(event_name, {})


def set_common_data_for_event(event_name, data):
    """Встановлює загальні дані для заходу"""
    event_common_data[event_name] = data


def remove_event_common_data(event_name):
    """Видаляє загальні дані заходу при видаленні заходу"""
    if event_name in event_common_data:
        del event_common_data[event_name]