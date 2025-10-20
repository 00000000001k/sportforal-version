# events.py

import customtkinter as ctk
from globals import document_blocks
from state_manager import save_current_state
from event_common_fields import create_common_fields_block, remove_event_common_data


def add_event(name, tabview, restore=False):
    if name in tabview._tab_dict:
        print(f"[WARN] Вкладка '{name}' вже існує — пропущено")
        return

    original_name = name
    counter = 1
    while name in tabview._tab_dict:
        name = f"{original_name}_{counter}"
        counter += 1

    tabview.add(name)
    tab_frame = tabview.tab(name)

    # ✅ Створюємо блок загальних полів зверху
    common_frame, common_entries = create_common_fields_block(tab_frame, name, tabview)

    # ✅ Скроллбар для договорів
    contracts_frame = ctk.CTkScrollableFrame(tab_frame)
    contracts_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # ✅ Зберігаємо посилання
    tab_frame.contracts_frame = contracts_frame
    tab_frame.common_frame = common_frame
    tab_frame.common_entries = common_entries

    if not restore:
        save_current_state(document_blocks, tabview)


def remove_tab(tab_name, tabview):
    """Видаляє вкладку та пов'язані з нею блоки"""
    global document_blocks

    # Видаляємо загальні дані заходу
    remove_event_common_data(tab_name)

    # Видаляємо вкладку
    tabview.delete(tab_name)

    # Видаляємо блоки договорів
    document_blocks[:] = [block for block in document_blocks if block["tab_name"] != tab_name]

    save_current_state(document_blocks, tabview)