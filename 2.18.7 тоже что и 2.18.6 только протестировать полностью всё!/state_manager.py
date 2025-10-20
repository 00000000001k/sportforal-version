# state_manager.py
# -*- coding: utf-8 -*-

import json
import os
from typing import Dict, Any, List
import tkinter.messagebox as messagebox
from globals import version

STATE_FILE = "app_state.json"


def save_current_state(document_blocks, tabview):
    """Зберігає поточний стан програми в новому форматі"""
    try:
        from event_common_fields import event_common_data

        state = {
            "version": version,
            "events": {},  # ИЗМЕНЕНИЕ: новая структура для событий
            "event_common_data": event_common_data,  # Сохраняем для совместимости
            "document_blocks": []
        }

        # НОВАЯ ЛОГИКА: Группируем данные по событиям
        events_data = {}

        # Собираем информацию о вкладках и связываем с событиями
        if hasattr(tabview, '_tab_dict') and tabview._tab_dict:
            for tab_name in tabview._tab_dict.keys():
                tab_frame = tabview.tab(tab_name)
                event_number = getattr(tab_frame, 'event_number', None)

                if event_number is not None:
                    # Инициализируем событие, если его еще нет
                    if event_number not in events_data:
                        events_data[event_number] = {
                            "name": tab_name,
                            "has_tab": True,
                            "is_current": tab_name == tabview.get(),
                            "common_data": event_common_data.get(event_number, {}),
                            "contracts": []
                        }

        # Сохраняем блоки договоров, привязывая к событиям
        for block in document_blocks:
            tab_name = block.get("tab_name", "")

            # Находим номер события для этой вкладки
            event_number = None
            if hasattr(tabview, '_tab_dict') and tab_name in tabview._tab_dict:
                tab_frame = tabview.tab(tab_name)
                event_number = getattr(tab_frame, 'event_number', None)

            if event_number is not None:
                # Инициализируем событие, если его еще нет
                if event_number not in events_data:
                    events_data[event_number] = {
                        "name": tab_name,
                        "has_tab": True,
                        "is_current": False,
                        "common_data": event_common_data.get(event_number, {}),
                        "contracts": []
                    }

                # Создаем данные договора
                contract_data = {
                    "path": block.get("path", ""),
                    "entries": {}
                }

                # Сохраняем данные из полей
                for field_key, entry_widget in block.get("entries", {}).items():
                    try:
                        contract_data["entries"][field_key] = entry_widget.get()
                    except:
                        contract_data["entries"][field_key] = ""

                events_data[event_number]["contracts"].append(contract_data)

            # Сохраняем в старом формате для совместимости
            block_data = {
                "path": block.get("path", ""),
                "tab_name": tab_name,
                "entries": {}
            }

            for field_key, entry_widget in block.get("entries", {}).items():
                try:
                    block_data["entries"][field_key] = entry_widget.get()
                except:
                    block_data["entries"][field_key] = ""

            state["document_blocks"].append(block_data)

        # Сохраняем события, которые есть в event_common_data, но не имеют вкладок
        for event_number, common_data in event_common_data.items():
            if event_number not in events_data:
                events_data[event_number] = {
                    "name": f"Захід #{event_number}",
                    "has_tab": False,
                    "is_current": False,
                    "common_data": common_data,
                    "contracts": []
                }

        state["events"] = events_data

        # Запись в файл
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)

        print(f"[INFO] Стан збережено: {len(events_data)} подій")
        print(f"[INFO] Події з вкладками: {[num for num, data in events_data.items() if data['has_tab']]}")

    except Exception as e:
        print(f"[ERROR] Помилка збереження стану: {e}")
        import traceback
        traceback.print_exc()


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

        # Відновлюємо вкладки
        current_tab = None
        restored_tabs = {}

        for tab_info in state.get("tabs", []):
            tab_name = tab_info["name"]
            event_number = tab_info.get("event_number", None)
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

        # Відновлюємо блоки договорів для кожної вкладки окремо
        for tab_name, tab_blocks in blocks_by_tab.items():
            if not tab_name or tab_name not in tabview._tab_dict:
                print(f"[WARN] Вкладка '{tab_name}' не знайдена, пропускаємо")
                continue

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
            save_current_state(document_blocks, tabview)
        except Exception as e:
            print(f"[ERROR] Помилка при збереженні стану: {e}")
        finally:
            root.destroy()

    # Прив'язуємо обробник до події закриття вікна
    root.protocol("WM_DELETE_WINDOW", on_closing)


def clear_saved_state():
    """Очищає збережений стан (для налагодження)"""
    try:
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
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
    """Отримує список існуючих подій з файлу стану"""
    try:
        if not os.path.exists(STATE_FILE):
            return {}

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        # Проверяем новый формат
        if "events" in state:
            return state["events"]

        # Миграция со старого формата
        return migrate_old_format(state)

    except Exception as e:
        print(f"[ERROR] Помилка читання стану: {e}")
        return {}


def migrate_old_format(old_state):
    """Міграція зі старого формату в новий"""
    print("[INFO] Виконується міграція зі старого формату...")

    events = {}
    event_counter = 1

    # Собираем данные из старого формата
    tabs = old_state.get("tabs", [])
    document_blocks = old_state.get("document_blocks", [])
    event_common_data = old_state.get("event_common_data", {})

    # Создаем события на основе вкладок
    for tab in tabs:
        tab_name = tab.get("name", "")
        event_number = tab.get("event_number", event_counter)

        events[event_number] = {
            "name": tab_name,
            "has_tab": True,
            "is_current": tab.get("is_current", False),
            "common_data": event_common_data.get(event_number, {}),
            "contracts": []
        }

        # Собираем контракты для этой вкладки
        for block in document_blocks:
            if block.get("tab_name") == tab_name:
                contract_data = {
                    "path": block.get("path", ""),
                    "entries": block.get("entries", {})
                }
                events[event_number]["contracts"].append(contract_data)

        event_counter += 1

    # Добавляем события без вкладок из event_common_data
    for event_num, common_data in event_common_data.items():
        if event_num not in events:
            events[event_num] = {
                "name": f"Захід #{event_num}",
                "has_tab": False,
                "is_current": False,
                "common_data": common_data,
                "contracts": []
            }

    print(f"[INFO] Міграція завершена: {len(events)} подій")

    # Сохраняем в новом формате
    try:
        new_state = {
            "version": old_state.get("version", version),
            "events": events,
            "event_common_data": event_common_data,
            "document_blocks": document_blocks  # Сохраняем для совместимости
        }

        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(new_state, f, indent=2, ensure_ascii=False)

        print("[INFO] Новий формат збережено")

    except Exception as e:
        print(f"[ERROR] Помилка збереження після міграції: {e}")

    return events


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
            set_common_data_for_event(event_number, common_data)

        # Створюємо захід
        from events import add_event
        add_event(event_name, tabview, restore=True, event_number=event_number)

        # Відновлюємо договори з нового формата
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

        # Якщо захід не мав вкладки, оновлюємо стан
        if not has_existing_tab:
            # Обновляем флаг has_tab в памяти
            existing_events[event_number]["has_tab"] = True
            existing_events[event_number]["name"] = event_name

        # Активуємо відновлену вкладку
        tabview.set(event_name)

        print(f"[INFO] Захід №{event_number} '{event_name}' відновлено успішно")
        return True

    except Exception as e:
        print(f"[ERROR] Помилка відновлення заходу {event_number}: {e}")
        import traceback
        traceback.print_exc()
        return False


def update_state_with_new_tab(tab_name, event_number):
    """Оновлює стан після створення нової вкладки"""
    try:
        existing_events = get_existing_events()

        if event_number in existing_events:
            existing_events[event_number]["has_tab"] = True
            existing_events[event_number]["name"] = tab_name

            # Сохраняем обновленное состояние
            if os.path.exists(STATE_FILE):
                with open(STATE_FILE, 'r', encoding='utf-8') as f:
                    state = json.load(f)

                state["events"] = existing_events

                with open(STATE_FILE, 'w', encoding='utf-8') as f:
                    json.dump(state, f, indent=2, ensure_ascii=False)

                print(f"[INFO] Оновлено стан для події #{event_number}")

    except Exception as e:
        print(f"[ERROR] Помилка оновлення стану: {e}")


def delete_event_from_state(event_number):
    """Видаляє захід зі збереженого стану"""
    try:
        if not os.path.exists(STATE_FILE):
            return True

        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)

        # Видаляємо з нового формату
        if "events" in state and event_number in state["events"]:
            event_name_to_delete = state["events"][event_number]["name"]
            del state["events"][event_number]
        else:
            event_name_to_delete = None

        # Знаходимо назву заходу за номером у старому форматі
        tabs_to_keep = []
        for tab_info in state.get("tabs", []):
            if tab_info.get("event_number") == event_number:
                if not event_name_to_delete:
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
        if str(event_number) in state.get("event_common_data", {}):
            del state["event_common_data"][str(event_number)]

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
            has_tab = "✓" if event_data.get('has_tab', False) else "✗"
            summary.append(f"№{event_num}: {name} ({contracts_count} договорів) [{has_tab}]")

        return "\n".join(summary)

    except Exception as e:
        return f"Помилка: {e}"


def get_events_without_tabs():
    """Отримує список подій без активних вкладок"""
    try:
        existing_events = get_existing_events()
        return {
            num: data for num, data in existing_events.items()
            if not data.get("has_tab", False)
        }
    except Exception as e:
        print(f"[ERROR] Помилка отримання подій: {e}")
        return {}


def restore_missing_tabs_from_common_data():
    """Восстанавливает отсутствующие вкладки на основе event_common_data"""
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
                "is_current": False,
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
    """Исправляет состояние и восстанавливает отсутствующие вкладки в интерфейсе"""
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