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
                state["tabs"].append({
                    "name": tab_name,
                    "is_current": tab_name == tabview.get()
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

        print(f"[INFO] Стан збережено: {len(state['tabs'])} вкладок, {len(state['document_blocks'])} договорів")

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

        print(
            f"[INFO] Завантажено стан: {len(state.get('tabs', []))} вкладок, {len(state.get('document_blocks', []))} договорів")
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

        # Відновлюємо загальні дані заходів
        if "event_common_data" in state:
            for event_name, data in state["event_common_data"].items():
                set_common_data_for_event(event_name, data)

        # Відновлюємо вкладки
        current_tab = None
        for tab_info in state.get("tabs", []):
            tab_name = tab_info["name"]
            add_event(tab_name, tabview, restore=True)
            if tab_info.get("is_current", False):
                current_tab = tab_name

        # Встановлюємо активну вкладку
        if current_tab and current_tab in tabview._tab_dict:
            tabview.set(current_tab)

        # Відновлюємо блоки договорів
        for block_data in state.get("document_blocks", []):
            tab_name = block_data.get("tab_name", "")
            template_path = block_data.get("path", "")

            if not tab_name or not template_path:
                continue

            # Перевіряємо, чи існує файл шаблону
            if not os.path.exists(template_path):
                print(f"[WARN] Шаблон не знайдено: {template_path}")
                continue

            # Знаходимо фрейм для договорів цієї вкладки
            if tab_name in tabview._tab_dict:
                tab_frame = tabview.tab(tab_name)
                if hasattr(tab_frame, 'contracts_frame'):
                    # Створюємо блок договору
                    create_document_fields_block(
                        tab_frame.contracts_frame,
                        tabview,
                        template_path
                    )

                    # Відновлюємо дані в полях (знаходимо останній створений блок)
                    from globals import document_blocks
                    if document_blocks:
                        last_block = document_blocks[-1]
                        entries_data = block_data.get("entries", {})

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

        print("[INFO] Стан програми відновлено успішно")

    except Exception as e:
        print(f"[ERROR] Помилка відновлення стану: {e}")
        messagebox.showerror("Помилка", f"Не вдалося повністю відновити стан програми:\n{e}")


def setup_auto_save(root, document_blocks, tabview):
    """Налаштовує автоматичне збереження при закритті програми"""

    def on_closing():
        """Обробник закриття програми"""
        try:
            print("[INFO] Програма закривається, збереження стану...")
            save_current_state(document_blocks, tabview)
            print("[INFO] Стан збережено успішно")
        except Exception as e:
            print(f"[ERROR] Помилка при збереженні стану: {e}")
        finally:
            root.destroy()

    # Прив'язуємо обробник до події закриття вікна
    root.protocol("WM_DELETE_WINDOW", on_closing)

    print("[INFO] Автоматичне збереження налаштовано")


def clear_saved_state():
    """Очищає збережений стан (для налагодження)"""
    try:
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
            print("[INFO] Збережений стан очищено")
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