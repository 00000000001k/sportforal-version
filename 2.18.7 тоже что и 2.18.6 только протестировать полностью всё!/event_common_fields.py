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

# ИСПРАВЛЕНИЕ: Теперь храним данные по номерам событий
# Глобальний словник для збереження загальних даних кожного заходу
# Структура: {event_number: {"name": "название", "fields": {"захід": "...", "дата": "..."}}}
event_common_data = {}


def update_common_fields_display(event_name, event_number, tabview):
    """Оновлює відображення загальних полів після відновлення"""
    if event_name not in tabview._tab_dict:
        return

    tab_frame = tabview.tab(event_name)
    if not hasattr(tab_frame, 'common_entries'):
        return

    common_entries = tab_frame.common_entries
    event_data = event_common_data.get(event_number, {}).get("fields", {})

    for field_key, entry_widget in common_entries.items():
        if field_key in event_data:
            try:
                entry_widget.set_text(event_data[field_key])
                print(
                    f"[INFO] Оновлено поле '{field_key}' для заходу #{event_number} '{event_name}': {event_data[field_key]}")
            except Exception as e:
                print(f"[ERROR] Помилка оновлення поля '{field_key}': {e}")


def create_common_fields_block(parent_frame, event_name, tabview, event_number):
    """
    Створює блок загальних полів для заходу

    Args:
        parent_frame: родительский фрейм
        event_name: название события
        tabview: объект TabView
        event_number: номер события (НОВЫЙ ОБЯЗАТЕЛЬНЫЙ ПАРАМЕТР)
    """
    # Основний фрейм для загальних полів
    common_frame = ctk.CTkFrame(parent_frame, border_width=2, border_color="green")
    common_frame.pack(pady=(5, 15), padx=5, fill="x")

    # Заголовок з кнопкою видалення
    header_frame = ctk.CTkFrame(common_frame, fg_color="transparent")
    header_frame.pack(fill="x", padx=10, pady=(10, 5))

    header_label = ctk.CTkLabel(header_frame,
                                text=f"📋 ЗАГАЛЬНІ ДАНІ ЗАХОДУ #{event_number}",
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

        # ИЗМЕНЕНИЕ: Загружаем по номеру события
        if event_number in event_common_data and field_key in event_common_data[event_number].get("fields", {}):
            entry.set_text(event_common_data[event_number]["fields"][field_key])

        # ИЗМЕНЕНИЕ: Прив'язуємо функцію оновлення при зміні (передаем event_number)
        entry.bind("<KeyRelease>", lambda e, field=field_key: update_common_field(event_name, event_number, field, e))
        entry.bind("<FocusOut>", lambda e, field=field_key: update_common_field(event_name, event_number, field, e))

    # ИЗМЕНЕНИЕ: Ініціалізуємо дані для заходу по номеру
    if event_number not in event_common_data:
        event_common_data[event_number] = {
            "name": event_name,
            "fields": {}
        }

    return common_frame, common_entries


def update_common_field(event_name, event_number, field_key, event):
    """
    Обновляет общее поле события

    Args:
        event_name: название события
        event_number: номер события (НОВЫЙ ПАРАМЕТР)
        field_key: ключ поля
        event: событие tkinter
    """
    try:
        # ИЗМЕНЕНИЕ: Работаем с номером события
        if event_number not in event_common_data:
            event_common_data[event_number] = {
                "name": event_name,
                "fields": {}
            }

        # Получаем значение из поля
        field_value = event.widget.get()

        # ИЗМЕНЕНИЕ: Сохраняем по номеру события
        event_common_data[event_number]["fields"][field_key] = field_value

        print(f"[DEBUG] Обновлено поле '{field_key}' для события #{event_number}: '{field_value}'")

    except Exception as e:
        print(f"[ERROR] Ошибка обновления общего поля: {e}")


def sync_common_fields_to_contracts(event_name, event_number):
    """Синхронізує загальні поля з усіма договорами заходу"""
    from globals import document_blocks

    if event_number not in event_common_data:
        return

    # Знаходимо всі блоки цього заходу
    event_blocks = [block for block in document_blocks if block.get("tab_name") == event_name]

    # Оновлюємо загільні поля в кожному блоці
    event_fields = event_common_data[event_number].get("fields", {})
    for block in event_blocks:
        entries = block.get("entries", {})
        for field_key in COMMON_FIELDS:
            if field_key in entries and field_key in event_fields:
                entry_widget = entries[field_key]
                # Зберігаємо поточний стан
                current_state = entry_widget.cget("state")
                # Тимчасово робимо поле доступним для редагування
                entry_widget.configure(state="normal")
                # Оновлюємо значення
                entry_widget.set_text(event_fields[field_key])
                # Повертаємо попередній стан
                entry_widget.configure(state=current_state)


def fill_common_fields_for_new_contract(event_name, event_number, contract_entries):
    """Заповнює загальні поля для нового договору"""
    if event_number not in event_common_data:
        return

    event_fields = event_common_data[event_number].get("fields", {})
    for field_key in COMMON_FIELDS:
        if field_key in contract_entries and field_key in event_fields:
            entry_widget = contract_entries[field_key]
            # Заповнюємо значенням із загальних даних
            entry_widget.set_text(event_fields[field_key])


def get_common_data_for_event(event_number):
    """Повертає загальні дані для заходу по номеру"""
    return event_common_data.get(event_number, {}).get("fields", {})


def set_common_data_for_event(event_number, event_name, data):
    """Встановлює загальні дані для заходу по номеру"""
    event_common_data[event_number] = {
        "name": event_name,
        "fields": data if isinstance(data, dict) else {}
    }


def remove_event_common_data(event_number):
    """Видаляє загальні дані заходу при видаленні заходу по номеру"""
    if event_number in event_common_data:
        del event_common_data[event_number]


# Функции для миграции старых данных
def migrate_old_event_data():
    """Мігрує дані зі старого формату (по іменах) в новий (по номерах)"""
    global event_common_data

    # Если есть данные в старом формате - конвертируем их
    old_data = {}
    new_data = {}

    for key, value in event_common_data.items():
        if isinstance(key, str):  # Старый формат - ключ это имя события
            old_data[key] = value
        else:  # Новый формат - ключ это номер события
            new_data[key] = value

    # Если есть старые данные, но нет информации о номерах - создаем с временными номерами
    if old_data and not new_data:
        print("[INFO] Міграція старих даних заходів...")
        temp_number = 1
        for event_name, event_data in old_data.items():
            new_data[temp_number] = {
                "name": event_name,
                "fields": event_data if isinstance(event_data, dict) else {}
            }
            temp_number += 1

        event_common_data = new_data
        print(f"[INFO] Перенесено {len(old_data)} заходів")


def get_event_number_by_name(event_name):
    """Знаходить номер заходу по його імені"""
    for event_number, event_data in event_common_data.items():
        if event_data.get("name") == event_name:
            return event_number
    return None