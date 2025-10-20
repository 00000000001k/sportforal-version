# state_manager.py
# -*- coding: utf-8 -*-

import json
import os
from typing import Dict, Any, List
import tkinter.messagebox as messagebox
from globals import version

STATE_FILE = "app_state.json"


def save_current_state(document_blocks, tabview):
    """Зберігає поточний стан програми"""
    try:
        from event_common_fields import event_common_data

        state = {
            "version": version,
            "tabs": [],
            "event_common_data": event_common_data,
            "document_blocks": []
        }

        # Збереження інформації про вкладки
        if hasattr(tabview, '_tab_dict') and tabview._tab_dict:
            for tab_name in tabview._tab_dict.keys():
                tab_frame = tabview.tab(tab_name)
                event_number = getattr(tab_frame, 'event_number', None)

                state["tabs"].append({
                    "name": tab_name,
                    "is_current": tab_name == tabview.get(),
                    "event_number": event_number
                })

        # Збереження блоків договорів
        for block in document_blocks:
            block_data = {
                "path": block.get("path", ""),
                "tab_name": block.get("tab_name", ""),
                "entries": {}
            }

            # Зберігаємо дані з полів
            for field_key, entry_widget in block.get("entries", {}).items():
                try:
                    block_data["entries"][field_key] = entry_widget.get()
                except:
                    block_data["entries"][field_key] = ""

            state["document_blocks"].append(block_data)

        # Запис у файл
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)

        #print(f"[INFO] Стан збережено: {len(state['tabs'])} вкладок, {len(state['document_blocks'])} договорів")
        #print(f"[INFO] Загальні дані заходів: {list(event_common_data.keys())}")

    except Exception as e:
        print(f"[ERROR] Помилка збереження стану: {e}")


def load_application_state():
    """Завантажує збережений стан програми"""
    if not os.path.exists(STATE_FILE):
        print("[INFO] Файл стану не знайдено, запуск з чистого листа")
        return None

    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        return state

    except Exception as e:
        print(f"[ERROR] Помилка завантаження стану: {e}")
        return None


def restore_application_state(state, tabview, main_frame):
    """Відновлює стан програми"""
    if not state:
        return

    try:
        from event_common_fields import event_common_data, set_common_data_for_event
        from events import add_event
        from document_block import create_document_fields_block
        from globals import document_blocks

        # ВАЖЛИВО: Очищуємо список блоків договорів перед відновленням
        document_blocks.clear()

        # Відновлюємо загальні дані заходів
        if "event_common_data" in state:
            for event_name, data in state["event_common_data"].items():
                set_common_data_for_event(event_name, data)
                #print(f"[INFO] Відновлено загальні дані для заходу '{event_name}': {data}")

        # Відновлюємо вкладки
        current_tab = None
        restored_tabs = {}

        for tab_info in state.get("tabs", []):
            tab_name = tab_info["name"]
            event_number = tab_info.get("event_number", None)
            #print(f"[INFO] Відновлюємо захід '{tab_name}' з номером {event_number}")
            add_event(tab_name, tabview, restore=True, event_number=event_number)
            restored_tabs[tab_name] = tab_info
            if tab_info.get("is_current", False):
                current_tab = tab_name

        # Встановлюємо активну вкладку
        if current_tab and current_tab in tabview._tab_dict:
            tabview.set(current_tab)

        # Групуємо блоки договорів по вкладках
        blocks_by_tab = {}
        for block_data in state.get("document_blocks", []):
            tab_name = block_data.get("tab_name", "")
            if tab_name not in blocks_by_tab:
                blocks_by_tab[tab_name] = []
            blocks_by_tab[tab_name].append(block_data)

        #print(f"[INFO] Блоки договорів по вкладках: {[(tab, len(blocks)) for tab, blocks in blocks_by_tab.items()]}")

        # Відновлюємо блоки договорів для кожної вкладки окремо
        for tab_name, tab_blocks in blocks_by_tab.items():
            if not tab_name or tab_name not in tabview._tab_dict:
                print(f"[WARN] Вкладка '{tab_name}' не знайдена, пропускаємо")
                continue

            #print(f"[INFO] Відновлюємо {len(tab_blocks)} договорів для вкладки '{tab_name}'")

            # Знаходимо фрейм для договорів цієї вкладки
            tab_frame = tabview.tab(tab_name)
            if not hasattr(tab_frame, 'contracts_frame'):
                print(f"[WARN] Фрейм для договорів не знайдено для вкладки '{tab_name}'")
                continue

            # Відновлюємо блоки договорів для цієї вкладки
            for block_data in tab_blocks:
                template_path = block_data.get("path", "")

                if not template_path:
                    print(f"[WARN] Порожній шлях до шаблону для вкладки '{tab_name}'")
                    continue

                # Перевіряємо, чи існує файл шаблону
                if not os.path.exists(template_path):
                    print(f"[WARN] Шаблон не знайдено: {template_path}")
                    continue

                # Встановлюємо активну вкладку перед створенням блоку
                tabview.set(tab_name)

                # Створюємо блок договору
                create_document_fields_block(
                    tab_frame.contracts_frame,
                    tabview,
                    template_path
                )

                # Відновлюємо дані в полях (знаходимо останній створений блок)
                if document_blocks:
                    last_block = document_blocks[-1]
                    entries_data = block_data.get("entries", {})

                    # Перевіряємо, що блок належить правильній вкладці
                    if last_block.get("tab_name") == tab_name:
                        #print(f"[INFO] Відновлюємо дані для {len(entries_data)} полів")

                        for field_key, saved_value in entries_data.items():
                            if field_key in last_block["entries"]:
                                entry_widget = last_block["entries"][field_key]
                                try:
                                    # Тимчасово робимо поле доступним
                                    current_state = entry_widget.cget("state")
                                    entry_widget.configure(state="normal")
                                    entry_widget.set_text(saved_value)
                                    entry_widget.configure(state=current_state)
                                except Exception as e:
                                    print(f"[WARN] Не вдалося відновити поле {field_key}: {e}")
                    else:
                        print(f"[ERROR] Блок належить вкладці '{last_block.get('tab_name')}', а не '{tab_name}'")

        # Встановлюємо активну вкладку в кінці
        if current_tab and current_tab in tabview._tab_dict:
            tabview.set(current_tab)

        #print(f"[INFO] Стан програми відновлено успішно. Всього блоків: {len(document_blocks)}")

        # Перевіряємо розподіл блоків по вкладках
        tab_distribution = {}
        for block in document_blocks:
            tab_name = block.get("tab_name", "невідомо")
            tab_distribution[tab_name] = tab_distribution.get(tab_name, 0) + 1

        #print(f"[INFO] Розподіл блоків по вкладках: {tab_distribution}")



    except Exception as e:
        print(f"[ERROR] Помилка відновлення стану: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Помилка", f"Не вдалося повністю відновити стан програми:\n{e}")


def setup_auto_save(root, document_blocks, tabview):
    """Налаштовує автоматичне збереження при закритті програми"""

    def on_closing():
        """Обробник закриття програми"""
        try:
            #print("[INFO] Програма закривається, збереження стану...")
            save_current_state(document_blocks, tabview)
            #print("[INFO] Стан збережено успішно")
        except Exception as e:
            print(f"[ERROR] Помилка при збереженні стану: {e}")
        finally:
            root.destroy()

    # Прив'язуємо обробник до події закриття вікна
    root.protocol("WM_DELETE_WINDOW", on_closing)

    #print("[INFO] Автоматичне збереження налаштовано")


def clear_saved_state():
    """Очищає збережений стан (для налагодження)"""
    try:
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
            #print("[INFO] Збережений стан очищено")
        else:
            print("[INFO] Файл стану не існує")
    except Exception as e:
        print(f"[ERROR] Помилка очищення стану: {e}")


def get_state_info():
    """Повертає інформацію про збережений стан"""
    if not os.path.exists(STATE_FILE):
        return "Файл стану не знайдено"

    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        info = []
        info.append(f"Версія: {state.get('version', 'невідома')}")
        info.append(f"Кількість вкладок: {len(state.get('tabs', []))}")
        info.append(f"Кількість договорів: {len(state.get('document_blocks', []))}")
        info.append(f"Заходи з загальними даними: {len(state.get('event_common_data', {}))}")

        return "\n".join(info)

    except Exception as e:
        return f"Помилка читання стану: {e}"