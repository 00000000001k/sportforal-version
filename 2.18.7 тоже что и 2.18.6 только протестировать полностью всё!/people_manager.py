# people_manager.py

import json
import os
import re

from globals import SPECIAL_ROLES, PEOPLE_CONFIG
from utils import get_executable_dir
from people_formatter import generate_replacements as external_generate_replacements


class PeopleManager:
    def __init__(self):
        self.settings_file = os.path.join(get_executable_dir(), "people_settings.json")
        self.selected_people = set()
        self.special_roles_selection = {}
        self._after_jobs = set()  # Отслеживание активных after() задач
        self.load_settings()

    def load_settings(self):
        """Завантажує збережені налаштування людей"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.selected_people = set(data.get('selected_people', []))
                    self.special_roles_selection = data.get('special_roles', {})
                    print(f"[INFO] Завантажено налаштування людей: {self.selected_people}")
            else:
                # Налаштування за замовчуванням
                self.special_roles_selection = {
                    "material_responsible": SPECIAL_ROLES["material_responsible"]["default"]
                }
                print("[INFO] Використовуються налаштування за замовчуванням")
        except Exception as e:
            print(f"[ERROR] Помилка завантаження налаштувань людей: {e}")
            self.selected_people = set()
            self.special_roles_selection = {
                "material_responsible": SPECIAL_ROLES["material_responsible"]["default"]
            }

    def save_settings(self):
        """Зберігає поточні налаштування людей"""
        try:
            data = {
                'selected_people': list(self.selected_people),
                'special_roles': self.special_roles_selection
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"[INFO] Збережено налаштування людей: {self.selected_people}")
        except Exception as e:
            print(f"[ERROR] Помилка збереження налаштувань людей: {e}")

    def cleanup_after_jobs(self):
        """Очищает активные after() задачи для предотвращения ошибок"""
        for job_id in self._after_jobs.copy():
            try:
                import tkinter as tk
                root = tk._default_root
                if root and root.winfo_exists():
                    root.after_cancel(job_id)
                self._after_jobs.discard(job_id)
            except Exception:
                pass

    def schedule_after(self, widget, delay, func):
        """Безопасный способ планирования after() задач"""
        try:
            if widget and hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                job_id = widget.after(delay, func)
                self._after_jobs.add(job_id)
                return job_id
        except Exception as e:
            print(f"[WARNING] Не удалось запланировать after() задачу: {e}")
        return None

    def toggle_person(self, person_id):
        """Перемикає вибір особи"""
        if person_id in self.selected_people:
            self.selected_people.remove(person_id)
        else:
            self.selected_people.add(person_id)
        self.save_settings()

    def is_person_selected(self, person_id):
        """Перевіряє чи обрана особа"""
        return person_id in self.selected_people

    def set_special_role(self, role_id, person_id):
        """Встановлює особу для спеціальної ролі"""
        if role_id in SPECIAL_ROLES:
            # Дозволяємо встановлювати None або "NONE" для скасування призначення
            if person_id in SPECIAL_ROLES[role_id]["options"] or person_id in [None, "NONE"]:
                # Якщо передається "NONE", зберігаємо як None
                if person_id == "NONE":
                    person_id = None
                self.special_roles_selection[role_id] = person_id
                self.save_settings()

    def get_special_role(self, role_id):
        """Отримує особу для спеціальної ролі"""
        return self.special_roles_selection.get(role_id, SPECIAL_ROLES.get(role_id, {}).get("default"))

    def get_person_data(self, person_id):
        """Отримує дані особи за ID"""
        return PEOPLE_CONFIG.get(person_id)

    def get_all_people(self):
        """Отримує всіх доступних людей"""
        return PEOPLE_CONFIG

    def get_selected_people_ordered(self):
        """Повертає обраних людей, відсортованих за рангом (від старшого до молодшого)"""
        selected_with_data = []
        for person_id in self.selected_people:
            person_data = PEOPLE_CONFIG.get(person_id)
            if person_data:
                selected_with_data.append((person_id, person_data))
        selected_with_data.sort(key=lambda x: x[1]['rank'])
        return selected_with_data

    def generate_people_list_text(self):
        """Генерує текст зі списком обраних людей у дательному відмінку через кому"""
        material_person_id = self.get_special_role("material_responsible")
        all_people_with_data = []

        for person_id in self.selected_people:
            person_data = PEOPLE_CONFIG.get(person_id)
            if person_data:
                all_people_with_data.append((person_id, person_data))

        # Перевіряємо, чи material_person_id не None і не пустий рядок
        if material_person_id and material_person_id not in self.selected_people:
            material_person_data = PEOPLE_CONFIG.get(material_person_id)
            if material_person_data:
                all_people_with_data.append((material_person_id, material_person_data))

        all_people_with_data.sort(key=lambda x: x[1]['rank'])

        print(f"[DEBUG] All people sorted by rank: {len(all_people_with_data)}")

        people_parts = []
        for person_id, person_data in all_people_with_data:
            people_part = f"{person_data['position_dative']} ({person_data['name_dative']})"
            people_parts.append(people_part)
            print(f"[DEBUG] Added person: {people_part}")

        result = ', '.join(people_parts) if people_parts else ""
        print(f"[DEBUG] Final people list text: '{result}' (length: {len(result)})")
        return result

    def detect_used_part_placeholders(self, text):
        """Определяет какие PART плейсхолдеры действительно используются в тексте"""
        used_parts = set()
        for i in range(1, 11):
            placeholder = f"{{{{SELECTED_PEOPLE_PART_{i}}}}}"
            if placeholder in text:
                used_parts.add(i)
        print(f"[DEBUG] Used PART placeholders in text: {sorted(used_parts)}")
        return used_parts

    def get_invisible_placeholder(self):
        """Возвращает невидимый символ для замены пустых плейсхолдеров"""
        return '\u200B'

    def clean_unused_placeholders(self, text):
        """Видаляє з тексту тільки ті плейсхолдери, які НЕ повинні були бути замінені"""
        print("[DEBUG] Starting clean_unused_placeholders")

        expected_replacements = self.generate_replacements()

        part_placeholders = set()
        for i in range(1, 6):
            part_placeholders.add(f"{{{{SELECTED_PEOPLE_PART_{i}}}}}")

        placeholders_with_content = set()
        for placeholder, replacement in expected_replacements.items():
            if replacement and replacement != self.get_invisible_placeholder():
                placeholders_with_content.add(placeholder)

        print(f"[DEBUG] PART placeholders (always processed): {part_placeholders}")
        print(f"[DEBUG] Placeholders with content: {placeholders_with_content}")

        pattern = r'\{\{[^}]+\}\}'
        placeholders_in_text = re.findall(pattern, text)
        print(f"[DEBUG] Found placeholders in text: {placeholders_in_text}")

        placeholders_to_remove = []
        for placeholder in placeholders_in_text:
            if placeholder in part_placeholders:
                print(f"[DEBUG] Keeping PART placeholder: {placeholder}")
                continue
            if placeholder in placeholders_with_content:
                print(f"[DEBUG] Keeping placeholder with content: {placeholder}")
                continue
            placeholders_to_remove.append(placeholder)

        print(f"[DEBUG] Placeholders to replace with invisible chars: {placeholders_to_remove}")

        invisible_char = self.get_invisible_placeholder()
        for placeholder in placeholders_to_remove:
            text = text.replace(placeholder, invisible_char)
            print(f"[DEBUG] Replaced unused placeholder with invisible char: {placeholder}")

        text = self.clean_text_formatting(text)
        return text

    def clean_text_formatting(self, text):
        """Очищає форматування тексту після видалення плейсхолдерів"""
        text = re.sub(r',\s*,', ',', text)
        text = re.sub(r'^\s*,\s*', '', text, flags=re.MULTILINE)
        text = re.sub(r',\s*$', '', text, flags=re.MULTILINE)
        text = re.sub(r'[ \t]+', ' ', text)
        return text.strip()

    # винесена функция в people_formatter
    def generate_replacements(self, text=None):
        return external_generate_replacements(
            selected_people = self.selected_people,
            get_special_role_func = self.get_special_role,
            generate_people_list_text_func = self.generate_people_list_text,
            detect_used_part_placeholders_func = self.detect_used_part_placeholders,
            invisible_char = self.get_invisible_placeholder(),
            text = text
        )


    def process_document_text(self, text):
        """Обробляє текст документа: виконує заміни та видаляє невикористані плейсхолдери"""
        print("[DEBUG] Starting process_document_text")
        original_text = text

        # Передаємо текст у generate_replacements для аналізу використовуваних плейсхолдерів
        replacements = self.generate_replacements(text)
        print(f"[DEBUG] Generated {len(replacements)} replacements")

        pattern = r'\{\{[^}]+\}\}'
        placeholders_in_text = re.findall(pattern, text)
        print(f"[DEBUG] Placeholders found in original text: {placeholders_in_text}")

        replacement_count = 0
        invisible_char = self.get_invisible_placeholder()

        for placeholder, replacement in replacements.items():
            if placeholder in text:
                old_text = text
                text = text.replace(placeholder, replacement)
                if old_text != text:
                    replacement_count += 1
                    if replacement and replacement != invisible_char:
                        print(f"[DEBUG] Successfully replaced {placeholder} with content (length: {len(replacement)})")
                    else:
                        print(f"[DEBUG] Successfully replaced {placeholder} with removed content")

        print(f"[DEBUG] Total successful replacements: {replacement_count}")

        remaining_placeholders = re.findall(pattern, text)
        if remaining_placeholders:
            print(f"[WARNING] Remaining unprocessed placeholders: {remaining_placeholders}")
            for placeholder in remaining_placeholders:
                text = text.replace(placeholder, invisible_char)
                print(f"[DEBUG] Replaced unprocessed placeholder with invisible char: {placeholder}")

        # Применяем новую функцию очистки, чтобы убрать лишние пустые строки
        text = self.advanced_cleanup_document(text)

        print(f"[DEBUG] Text processing completed. Original length: {len(original_text)}, Final length: {len(text)}")
        return text

    def advanced_cleanup_document(self, text):
        import re
        print("[DEBUG] Начинается advanced_cleanup_document")
        print("[DEBUG] Исходный текст:")
        print(text)

        # 1. Удаляем невидимые символы
        invisible_char = self.get_invisible_placeholder()
        text = text.replace(invisible_char, '')
        print("[DEBUG] После удаления невидимых символов:")
        print(text)

        # 2. Нормализуем переносы строк (CRLF/CR заменяем на LF)
        text = re.sub(r'\r\n|\r', '\n', text)
        print("[DEBUG] После нормализации переносов:")
        print(text)

        # 3. Многократно заменяем последовательности из 3 и более переносов (и пробельных символов) на ровно 2 переноса
        pattern = re.compile(r'(?:\n\s*){3,}')
        iteration = 1
        while pattern.search(text):
            print(f"[DEBUG] Итерация {iteration}: Найдена последовательность переносов, заменяем на '\\n\\n'")
            text = pattern.sub("\n\n", text)
            print("[DEBUG] Промежуточный результат:")
            print(text)
            iteration += 1

        # 4. Обрезка пробелов в начале и в конце
        text = text.strip()
        print("[DEBUG] Финальный очищенный текст:")
        print(text)

        return text

    def get_replacements_for_removal(self):
        """Генерує словник для видалення необраних людей з документа"""
        replacements = {}
        # Замість невидимого символа використовуємо порожній рядок (тобто – видаляємо блок)
        for person_id, person_data in PEOPLE_CONFIG.items():
            if person_id not in self.selected_people:
                replacements[person_data['placeholders']['full_block']] = ""
                replacements[person_data['placeholders']['name_only']] = ""
                replacements[person_data['placeholders']['position_only']] = ""
        return replacements

    def get_selected_count(self):
        """Повертає кількість обраних людей"""
        return len(self.selected_people)

    def get_summary(self):
        """Повертає короткий опис обраних людей"""
        ordered_people = self.get_selected_people_ordered()
        material_person_id = self.get_special_role("material_responsible")
        material_person = PEOPLE_CONFIG.get(material_person_id)
        material_name = material_person['name'] if material_person else "Не вибрано"

        if not ordered_people and not material_person:
            return "Жодна особа не обрана"

        summary = f"Обрано осіб: {len(ordered_people)} (за рангом)\n"
        for i, (_, person_data) in enumerate(ordered_people, 1):
            summary += f"{i}. {person_data['display_name']}\n"

        summary += f"\nМатвідповідальна: {material_name}"

        return summary.rstrip()

    # Методи для сумісності з новою системою
    def add_person(self, person_data):
        """Додає нову людину (заглушка для сумісності)"""
        pass

    def update_person(self, index, person_data):
        """Оновлює дані людини (заглушка для сумісності)"""
        pass

    def delete_person(self, index):
        """Видаляє людину (заглушка для сумісності)"""
        pass

    def get_people(self):
        """Повертає список всіх людей у форматі для нової системи"""
        people_list = []
        for person_id, person_data in PEOPLE_CONFIG.items():
            people_list.append({
                'ПІБ': person_data['name'],
                'посада': person_data['position'],
                'id': person_id,
                'rank': person_data['rank']
            })
        people_list.sort(key=lambda x: x['rank'])
        return people_list

    def get_person(self, index):
        """Повертає дані людини за індексом у форматі для нової системи"""
        people_list = self.get_people()
        if 0 <= index < len(people_list):
            return people_list[index]
        return None

    def get_person_count(self):
        """Повертає кількість людей"""
        return len(PEOPLE_CONFIG)

    def debug_test(self):
        """Тестовий метод для діагностики"""
        print("=== DEBUG TEST ===")
        print(f"Selected people: {self.selected_people}")
        print(f"Special roles: {self.special_roles_selection}")
        inv = self.get_invisible_placeholder()
        print(f"Invisible char: '{inv}' (Unicode: U+{ord(inv):04X})")

        people_list = self.generate_people_list_text()
        print(f"Generated people list: '{people_list}'")

        replacements = self.generate_replacements()
        print(f"SELECTED_PEOPLE_LIST replacement: '{replacements.get('{{SELECTED_PEOPLE_LIST}}', 'NOT FOUND')}'")

        invisible_char = self.get_invisible_placeholder()
        real_content = {k: v for k, v in replacements.items() if v and v != invisible_char}
        for placeholder, replacement in real_content.items():
            snippet = replacement[:50] + ("..." if len(replacement) > 50 else "")
            print(f"  {placeholder}: '{snippet}'")

        invisible_count = sum(1 for v in replacements.values() if v == invisible_char)
        print(f"Invisible char replacements: {invisible_count}")
        print("=== END DEBUG ===")
        return replacements

    def __del__(self):
        """Деструктор для очистки after() задач"""
        try:
            self.cleanup_after_jobs()
        except Exception:
            pass


# Глобальний екземпляр менеджера людей
people_manager = PeopleManager()

if __name__ == "__main__":
    # Тестування, якщо файл запускається напряму
    pm = PeopleManager()
    pm.selected_people.add("basai")
    pm.selected_people.add("mokina")
    pm.debug_test()
