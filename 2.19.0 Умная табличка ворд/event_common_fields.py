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


def update_common_fields_display(event_name, tabview):
    """Оновлює відображення загальних полів після відновлення"""
    if event_name not in tabview._tab_dict:
        return

    tab_frame = tabview.tab(event_name)
    if not hasattr(tab_frame, 'common_entries'):
        return

    common_entries = tab_frame.common_entries
    event_data = event_common_data.get(event_name, {})

    for field_key, entry_widget in common_entries.items():
        if field_key in event_data:
            try:
                entry_widget.set_text(event_data[field_key])
                print(f"[INFO] Оновлено поле '{field_key}' для заходу '{event_name}': {event_data[field_key]}")
            except Exception as e:
                print(f"[ERROR] Помилка оновлення поля '{field_key}': {e}")


def create_common_fields_block(parent_frame, event_name, tabview=None, event_number=None):
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

        # Прив'язуємо функцію оновлення при зміні (ЗМІНЕНО - передаємо event_number)
        entry.bind("<KeyRelease>", lambda e, field=field_key: update_common_field(event_name, field, e, event_number))
        entry.bind("<FocusOut>", lambda e, field=field_key: update_common_field(event_name, field, e, event_number))

    # Ініціалізуємо дані для заходу, якщо їх немає
    if event_name not in event_common_data:
        event_common_data[event_name] = {}

    return common_frame, common_entries


def update_common_field(event_name, field_key, event, event_number=None):
    """Оновлює загальне поле та синхронізує його з усіма договорами заходу"""
    from globals import document_blocks

    # Отримуємо нове значення
    new_value = event.widget.get()

    # Зберігаємо у глобальних даних
    if event_name not in event_common_data:
        event_common_data[event_name] = {}
    event_common_data[event_name][field_key] = new_value

    # Зберігаємо в файл з правильним event_number
    set_common_data_for_event(event_name, event_common_data[event_name], event_number=event_number)

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


def set_common_data_for_event(event_name, data, event_number=None):
    """Встановлює загальні дані для заходу

    Args:
        event_name: назва заходу
        data: дані заходу
        event_number: номер заходу (зберігається як original_event_number)
    """
    # Якщо передано номер заходу, зберігаємо його
    if event_number is not None:
        data["original_event_number"] = event_number

    event_common_data[event_name] = data


def remove_event_common_data(event_name):
    """Видаляє загальні дані заходу при видаленні заходу"""
    if event_name in event_common_data:
        del event_common_data[event_name]