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


def get_existing_events():
    """Повертає словник існуючих заходів з їх інформацією"""
    try:
        if not os.path.exists(STATE_FILE):
            return {}

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        existing_events = {}

        # Отримуємо інформацію про вкладки
        tabs_info = state.get("tabs", [])
        document_blocks = state.get("document_blocks", [])
        event_common_data = state.get("event_common_data", {})

        # Групуємо договори по заходах
        contracts_by_event = {}
        for block in document_blocks:
            tab_name = block.get("tab_name", "")
            if tab_name:
                if tab_name not in contracts_by_event:
                    contracts_by_event[tab_name] = []
                contracts_by_event[tab_name].append(block)

        # Створюємо словник заходів з існуючих вкладок
        for tab_info in tabs_info:
            tab_name = tab_info.get("name", "")
            event_number = tab_info.get("event_number")

            if event_number is not None:
                existing_events[event_number] = {
                    "name": tab_name,
                    "contracts": contracts_by_event.get(tab_name, []),
                    "common_data": event_common_data.get(tab_name, {}),
                    "tab_info": tab_info,
                    "has_tab": True  # Позначаємо, що вкладка існує
                }

        # ДОДАЄМО ЗАХОДИ З event_common_data, які не мають вкладок
        # Отримуємо імена заходів, які вже мають вкладки
        events_with_tabs = {event_data["name"] for event_data in existing_events.values()}

        # Додаємо заходи з common_data, які не мають вкладок
        for event_name, common_data in event_common_data.items():
            if event_name not in events_with_tabs:
                # ВИПРАВЛЕННЯ: Використовуємо збережений original_event_number
                original_number = common_data.get("original_event_number")

                if original_number is not None:
                    # Використовуємо оригінальний номер
                    event_number = original_number
                else:
                    # Якщо з якихось причин номер не збережений, використовуємо резервну логіку
                    max_event_number = max(existing_events.keys()) if existing_events else 0
                    event_number = max_event_number + 1
                    print(
                        f"[WARN] Не знайдено original_event_number для заходу '{event_name}', використовуємо {event_number}")

                existing_events[event_number] = {
                    "name": event_name,
                    "contracts": contracts_by_event.get(event_name, []),
                    "common_data": common_data,
                    "tab_info": None,  # Немає вкладки
                    "has_tab": False  # Позначаємо, що вкладка не існує
                }

        return existing_events

    except Exception as e:
        print(f"[ERROR] Помилка отримання існуючих заходів: {e}")
        return {}

def restore_single_event(event_number, tabview):
    """Відновлює один конкретний захід за номером"""
    try:
        existing_events = get_existing_events()

        if event_number not in existing_events:
            print(f"[ERROR] Захід з номером {event_number} не знайдено")
            return False

        event_data = existing_events[event_number]
        event_name = event_data["name"]
        has_existing_tab = event_data.get("has_tab", False)

        # Перевіряємо, чи не існує вже така вкладка в інтерфейсі
        if hasattr(tabview, '_tab_dict') and event_name in tabview._tab_dict:
            response = messagebox.askyesno(
                "Захід вже існує",
                f"Захід '{event_name}' вже відкритий.\nЗамінити його?"
            )
            if response:
                # Видаляємо існуючу вкладку
                from events import remove_tab
                from globals import document_blocks

                # Видаляємо блоки цього заходу
                document_blocks[:] = [block for block in document_blocks if block.get("tab_name") != event_name]

                # Видаляємо вкладку
                tabview.delete(event_name)
            else:
                return False

        # Відновлюємо загальні дані заходу
        from event_common_fields import set_common_data_for_event
        common_data = event_data.get("common_data", {})
        if common_data:
            set_common_data_for_event(event_name, common_data)

        # Створюємо захід
        from events import add_event
        add_event(event_name, tabview, restore=True, event_number=event_number)

        # Відновлюємо договори
        contracts = event_data.get("contracts", [])
        if contracts:
            tab_frame = tabview.tab(event_name)
            if hasattr(tab_frame, 'contracts_frame'):
                from document_block import create_document_fields_block
                from globals import document_blocks

                for contract_data in contracts:
                    template_path = contract_data.get("path", "")

                    if not template_path or not os.path.exists(template_path):
                        print(f"[WARN] Шаблон не знайдено: {template_path}")
                        continue

                    # Встановлюємо активну вкладку
                    tabview.set(event_name)

                    # Створюємо блок договору
                    create_document_fields_block(
                        tab_frame.contracts_frame,
                        tabview,
                        template_path
                    )

                    # Відновлюємо дані в полях
                    if document_blocks:
                        last_block = document_blocks[-1]
                        entries_data = contract_data.get("entries", {})

                        if last_block.get("tab_name") == event_name:
                            for field_key, saved_value in entries_data.items():
                                if field_key in last_block["entries"]:
                                    entry_widget = last_block["entries"][field_key]
                                    try:
                                        current_state = entry_widget.cget("state")
                                        entry_widget.configure(state="normal")
                                        entry_widget.set_text(saved_value)
                                        entry_widget.configure(state=current_state)
                                    except Exception as e:
                                        print(f"[WARN] Не вдалося відновити поле {field_key}: {e}")

        # Якщо захід не мав вкладки, додаємо її в збережений стан
        if not has_existing_tab:
            update_state_with_new_tab(event_name, event_number)

        # Активуємо відновлену вкладку
        tabview.set(event_name)

        print(f"[INFO] Захід №{event_number} '{event_name}' відновлено успішно")
        return True

    except Exception as e:
        print(f"[ERROR] Помилка відновлення заходу {event_number}: {e}")
        import traceback
        traceback.print_exc()
        return False


def update_state_with_new_tab(event_name, event_number):
    """Додає нову вкладку в збережений стан"""
    try:
        if not os.path.exists(STATE_FILE):
            return

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        # Перевіряємо, чи вже є така вкладка
        existing_tab_names = {tab.get("name") for tab in state.get("tabs", [])}

        if event_name not in existing_tab_names:
            new_tab = {
                "name": event_name,
                "is_current": True,  # Робимо активною
                "event_number": event_number
            }

            # Знімаємо is_current з інших вкладок
            for tab in state.get("tabs", []):
                tab["is_current"] = False

            state.setdefault("tabs", []).append(new_tab)

            # Зберігаємо оновлений стан
            with open(STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=2, ensure_ascii=False)

            print(f"[INFO] Додано вкладку '{event_name}' в збережений стан")

    except Exception as e:
        print(f"[ERROR] Помилка оновлення стану з новою вкладкою: {e}")


def delete_event_from_state(event_number):
    """Видаляє захід зі збереженого стану"""
    try:
        if not os.path.exists(STATE_FILE):
            return True

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        # Знаходимо назву заходу за номером
        event_name_to_delete = None
        tabs_to_keep = []

        for tab_info in state.get("tabs", []):
            if tab_info.get("event_number") == event_number:
                event_name_to_delete = tab_info.get("name")
            else:
                tabs_to_keep.append(tab_info)

        if not event_name_to_delete:
            return True  # Захід не знайдено, нічого видаляти

        # Оновлюємо список вкладок
        state["tabs"] = tabs_to_keep

        # Видаляємо договори цього заходу
        state["document_blocks"] = [
            block for block in state.get("document_blocks", [])
            if block.get("tab_name") != event_name_to_delete
        ]

        # Видаляємо загальні дані заходу
        if event_name_to_delete in state.get("event_common_data", {}):
            del state["event_common_data"][event_name_to_delete]

        # Зберігаємо оновлений стан
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)

        print(f"[INFO] Захід №{event_number} '{event_name_to_delete}' видалено зі стану")
        return True

    except Exception as e:
        print(f"[ERROR] Помилка видалення заходу зі стану: {e}")
        return False


def get_events_summary():
    """Повертає короткий огляд всіх збережених заходів"""
    try:
        existing_events = get_existing_events()

        if not existing_events:
            return "Немає збережених заходів"

        summary = []
        for event_num, event_data in sorted(existing_events.items()):
            name = event_data.get('name', 'Без назви')
            contracts_count = len(event_data.get('contracts', []))
            summary.append(f"№{event_num}: {name} ({contracts_count} договорів)")

        return "\n".join(summary)

    except Exception as e:
        return f"Помилка: {e}"


def restore_missing_tabs_from_common_data():
    """
    Восстанавливает отсутствующие вкладки на основе event_common_data
    """
    try:
        if not os.path.exists(STATE_FILE):
            print("[INFO] Файл состояния не найден")
            return False

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        # Получаем существующие вкладки
        existing_tabs = {tab["name"] for tab in state.get("tabs", [])}

        # Получаем все события из common_data
        common_data_events = set(state.get("event_common_data", {}).keys())

        print(f"[INFO] Существующие вкладки: {existing_tabs}")
        print(f"[INFO] События в common_data: {common_data_events}")

        # Находим отсутствующие вкладки
        missing_events = common_data_events - existing_tabs

        if not missing_events:
            print("[INFO] Все события уже имеют соответствующие вкладки")
            return False

        print(f"[INFO] Найдены отсутствующие вкладки: {missing_events}")

        # Определяем следующий доступный номер события
        existing_numbers = {tab.get("event_number") for tab in state.get("tabs", []) if
                            tab.get("event_number") is not None}
        next_number = 1
        while next_number in existing_numbers:
            next_number += 1

        # Добавляем отсутствующие вкладки
        tabs_added = 0
        for event_name in missing_events:
            new_tab = {
                "name": event_name,
                "is_current": False,  # Не делаем активной по умолчанию
                "event_number": next_number
            }

            state["tabs"].append(new_tab)
            print(f"[INFO] Добавлена вкладка: {event_name} с номером {next_number}")

            next_number += 1
            tabs_added += 1

        # Сохраняем обновленное состояние
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)

        print(f"[INFO] Добавлено {tabs_added} отсутствующих вкладок")
        return True

    except Exception as e:
        print(f"[ERROR] Ошибка восстановления отсутствующих вкладок: {e}")
        return False


def fix_state_and_restore_missing(tabview):
    """
    Исправляет состояние и восстанавливает отсутствующие вкладки в интерфейсе
    """
    try:
        # Сначала исправляем JSON файл
        if not restore_missing_tabs_from_common_data():
            print("[INFO] Нет отсутствующих вкладок для восстановления")
            return False

        # Теперь восстанавливаем интерфейс
        if not os.path.exists(STATE_FILE):
            return False

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        from event_common_fields import set_common_data_for_event
        from events import add_event

        # Получаем существующие вкладки в интерфейсе
        existing_interface_tabs = set()
        if hasattr(tabview, '_tab_dict') and tabview._tab_dict:
            existing_interface_tabs = set(tabview._tab_dict.keys())

        print(f"[INFO] Вкладки в интерфейсе: {existing_interface_tabs}")

        # Восстанавливаем недостающие вкладки
        restored_count = 0
        for tab_info in state.get("tabs", []):
            tab_name = tab_info["name"]
            event_number = tab_info.get("event_number")

            if tab_name not in existing_interface_tabs:
                print(f"[INFO] Восстанавливаем вкладку: {tab_name} (№{event_number})")

                # Восстанавливаем общие данные события
                common_data = state.get("event_common_data", {}).get(tab_name, {})
                if common_data:
                    set_common_data_for_event(tab_name, common_data)

                # Создаем вкладку
                add_event(tab_name, tabview, restore=True, event_number=event_number)
                restored_count += 1

        print(f"[INFO] Восстановлено {restored_count} вкладок в интерфейсе")
        return restored_count > 0

    except Exception as e:
        print(f"[ERROR] Ошибка исправления состояния: {e}")
        import traceback
        traceback.print_exc()
        return False
