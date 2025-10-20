# events.py

import customtkinter as ctk
from globals import document_blocks
from state_manager import save_current_state


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

    # ✅ Скроллбар на всю вкладку
    contracts_frame = ctk.CTkScrollableFrame(tab_frame)  # исправлено
    contracts_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # ✅ Сохраняем этот scrollable frame как свойство
    tab_frame.contracts_frame = contracts_frame

    if not restore:
        save_current_state(document_blocks, tabview)

def remove_tab(tab_name, tabview):
    """Видаляє вкладку та повʼязані з нею блоки"""
    global document_blocks
    tabview.delete(tab_name)
    document_blocks[:] = [block for block in document_blocks if block["tab_name"] != tab_name]
    save_current_state(document_blocks, tabview)
