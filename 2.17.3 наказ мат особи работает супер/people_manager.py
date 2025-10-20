# people_manager.py - ВИПРАВЛЕНА ВЕРСІЯ з активною діагностикою

import json
import os
import re

from globals import SPECIAL_ROLES, PEOPLE_CONFIG
from utils import get_executable_dir


class PeopleManager:
    def __init__(self):
        self.settings_file = os.path.join(get_executable_dir(), "people_settings.json")
        self.selected_people = set()
        self.special_roles_selection = {}
        self._after_jobs = set()
        # ДОДАЄМО ФЛАГ ДЛЯ ДІАГНОСТИКИ
        self.debug_mode = True
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
        """Очищает активные after() задачи"""
        for job_id in self._after_jobs.copy():
            try:
                import tkinter as tk
                root = tk._default_root
                if root and root.winfo_exists():
                    root.after_cancel(job_id)
                self._after_jobs.discard(job_id)
            except Exception:
                pass

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
        if role_id in SPECIAL_ROLES and person_id in SPECIAL_ROLES[role_id]["options"]:
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
        """Повертає обраних людей, відсортованих за рангом"""
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

        # Добавляем всех обранных людей
        for person_id in self.selected_people:
            person_data = PEOPLE_CONFIG.get(person_id)
            if person_data:
                all_people_with_data.append((person_id, person_data))

        # Добавляем матеріально-відповідальну особу, если она не в списке обранных
        if material_person_id and material_person_id not in self.selected_people:
            material_person_data = PEOPLE_CONFIG.get(material_person_id)
            if material_person_data:
                all_people_with_data.append((material_person_id, material_person_data))

        # Сортируем по рангу
        all_people_with_data.sort(key=lambda x: x[1]['rank'])

        people_parts = []
        for person_id, person_data in all_people_with_data:
            people_part = f"{person_data['position_dative']} ({person_data['name_dative']})"
            people_parts.append(people_part)

        return ', '.join(people_parts) if people_parts else ""

    def detect_used_part_placeholders(self, text):
        """Определяет какие PART плейсхолдеры действительно используются в тексте"""
        used_parts = set()
        for i in range(1, 11):
            placeholder = f"{{{{SELECTED_PEOPLE_PART_{i}}}}}"
            if placeholder in text:
                used_parts.add(i)
        if self.debug_mode:
            print(f"[DEBUG] Used PART placeholders in text: {sorted(used_parts)}")
        return used_parts

    def detect_used_individual_placeholders(self, text):
        """Определяет какие индивидуальные плейсхолдеры используются в тексте"""
        used_placeholders = set()

        # Проверяем все индивидуальные плейсхолдеры
        for person_id, person_data in PEOPLE_CONFIG.items():
            for placeholder_type, placeholder in person_data['placeholders'].items():
                if placeholder in text:
                    used_placeholders.add(placeholder)

        # Проверяем специальные роли
        for role_config in SPECIAL_ROLES.values():
            if role_config['placeholder'] in text:
                used_placeholders.add(role_config['placeholder'])

        if self.debug_mode:
            print(f"[DEBUG] Used individual placeholders in text: {len(used_placeholders)}")
        return used_placeholders

    def generate_individual_replacements(self, text):
        """
        ПЕРВАЯ ЛОГИКА: Генерує заміни для індивідуальних плейсхолдерів
        Використовує оригінальну логіку з першого файлу
        """
        if self.debug_mode:
            print("[DEBUG] Generating individual replacements (первая логика)")
        replacements = {}

        # Определяем какие индивидуальные плейсхолдеры есть в тексте
        used_individual_placeholders = self.detect_used_individual_placeholders(text)

        # Обрабатываем только те плейсхолдеры, которые есть в тексте
        for person_id, person_data in PEOPLE_CONFIG.items():
            person_placeholders = person_data['placeholders']

            # Проверяем каждый тип плейсхолдера для этой персоны
            for placeholder_type, placeholder in person_placeholders.items():
                if placeholder in used_individual_placeholders:
                    # Если персона выбрана - даем контент
                    if person_id in self.selected_people:
                        if placeholder_type == 'full_block':
                            # Полный блок с оригинальным форматированием
                            content = (
                                f"{person_data['position']}\r\n"
                                f"\r\n"
                                f"1___1________________2025 року \t\t\t\t{person_data['name']}"
                            )
                        elif placeholder_type == 'name_only':
                            content = person_data['name']
                        elif placeholder_type == 'position_only':
                            content = person_data['position']
                        else:
                            content = ""

                        replacements[placeholder] = content
                        if self.debug_mode:
                            print(f"[DEBUG] Individual: {placeholder} -> content (selected)")
                    else:
                        # Если персона НЕ выбрана - удаляем (пустая строка)
                        replacements[placeholder] = ""
                        if self.debug_mode:
                            print(f"[DEBUG] Individual: {placeholder} -> empty (not selected)")

        # Обрабатываем специальные роли
        for role_id, role_config in SPECIAL_ROLES.items():
            if role_config['placeholder'] in used_individual_placeholders:
                selected_person_id = self.get_special_role(role_id)
                person_data = PEOPLE_CONFIG.get(selected_person_id)

                if person_data:
                    if role_id == "material_responsible":
                        content = (
                            f"{person_data['position']}\r\n"
                            f"\r\n"
                            f"1___1________________2025 року \t\t\t\t{person_data['name']}"
                        )
                        replacements[role_config['placeholder']] = content
                        if self.debug_mode:
                            print(f"[DEBUG] Special role: {role_config['placeholder']} -> content")
                else:
                    replacements[role_config['placeholder']] = ""
                    if self.debug_mode:
                        print(f"[DEBUG] Special role: {role_config['placeholder']} -> empty")

        return replacements

    def generate_parts_replacements(self, text):
        """
        ВТОРАЯ ЛОГИКА: Генерує заміни для PART плейсхолдерів
        Використовує покращену логіку з другого файлу
        """
        if self.debug_mode:
            print("[DEBUG] Generating parts replacements (вторая логика)")
        replacements = {}

        # Определяем какие PART плейсхолдеры есть в тексте
        used_part_numbers = self.detect_used_part_placeholders(text)

        # Если нет PART плейсхолдеров в тексте - не обрабатываем
        if not used_part_numbers:
            if self.debug_mode:
                print("[DEBUG] No PART placeholders found in text")
            return replacements

        # Генерируем текст списка людей
        people_list_text = self.generate_people_list_text()

        if not people_list_text:
            # Если нет людей - все PART плейсхолдеры остаются пустыми
            for part_num in used_part_numbers:
                placeholder = f"{{{{SELECTED_PEOPLE_PART_{part_num}}}}}"
                replacements[placeholder] = ""
                if self.debug_mode:
                    print(f"[DEBUG] Parts: {placeholder} -> empty (no people)")

            # Обрабатываем связанные плейсхолдеры
            if "{{SELECTED_PEOPLE_LIST}}" in text:
                replacements["{{SELECTED_PEOPLE_LIST}}"] = ""
            if "{{SELECTED_PEOPLE_PARTS_COUNT}}" in text:
                replacements["{{SELECTED_PEOPLE_PARTS_COUNT}}"] = "0"

            return replacements

        # Разбиваем текст на части если он длинный
        if len(people_list_text) > 150:
            if self.debug_mode:
                print(f"[DEBUG] Long text ({len(people_list_text)} chars), splitting into parts")

            parts = people_list_text.split(', ')
            chunks = []
            current_chunk = ""

            for part in parts:
                test_chunk = current_chunk + (", " if current_chunk else "") + part
                if len(test_chunk) <= 150:
                    current_chunk = test_chunk
                else:
                    if current_chunk:
                        chunks.append(current_chunk)
                    current_chunk = part

            if current_chunk:
                chunks.append(current_chunk)

            if self.debug_mode:
                print(f"[DEBUG] Split into {len(chunks)} chunks")

            # Заполняем только используемые PART плейсхолдеры
            filled_parts = 0
            for i in range(len(chunks)):
                part_number = i + 1
                if part_number in used_part_numbers:
                    placeholder = f"{{{{SELECTED_PEOPLE_PART_{part_number}}}}}"

                    # Для первой части - без префикса
                    if i == 0:
                        replacements[placeholder] = chunks[i]
                    else:
                        # Для следующих частей - добавляем запятую
                        replacements[placeholder] = f", {chunks[i]}"

                    filled_parts += 1
                    if self.debug_mode:
                        print(f"[DEBUG] Parts: {placeholder} -> content (chunk {i + 1})")

            # Пустые значения для неиспользуемых PART плейсхолдеров
            for part_num in used_part_numbers:
                placeholder = f"{{{{SELECTED_PEOPLE_PART_{part_num}}}}}"
                if placeholder not in replacements:
                    replacements[placeholder] = ""
                    if self.debug_mode:
                        print(f"[DEBUG] Parts: {placeholder} -> empty (unused)")

        else:
            # Короткий текст - помещаем в первую часть
            filled_parts = 0
            for part_num in used_part_numbers:
                placeholder = f"{{{{SELECTED_PEOPLE_PART_{part_num}}}}}"
                if part_num == 1:
                    replacements[placeholder] = people_list_text
                    filled_parts = 1
                    if self.debug_mode:
                        print(f"[DEBUG] Parts: {placeholder} -> full text")
                else:
                    replacements[placeholder] = ""
                    if self.debug_mode:
                        print(f"[DEBUG] Parts: {placeholder} -> empty")

        # Обрабатываем связанные плейсхолдеры
        if "{{SELECTED_PEOPLE_LIST}}" in text:
            replacements["{{SELECTED_PEOPLE_LIST}}"] = people_list_text

        if "{{SELECTED_PEOPLE_PARTS_COUNT}}" in text:
            replacements["{{SELECTED_PEOPLE_PARTS_COUNT}}"] = str(filled_parts)

        return replacements

    # ГОЛОВНИЙ МЕТОД generate_replacements - ВИПРАВЛЕНА ВЕРСІЯ
    def generate_replacements(self, text=None):
        """
        ГОЛОВНИЙ МЕТОД: Генерує всі заміни з детальною діагностикою
        """
        print("=" * 50)
        print("🔍 ПОЧАТОК ГЕНЕРАЦІЇ ЗАМІННИКІВ")
        print("=" * 50)

        # КРОК 1: Перевірка вхідного тексту
        if text is None:
            print("❌ КРИТИЧНА ПОМИЛКА: Параметр text є None!")
            print("💡 Перевірте виклик методу generate_replacements()")
            import traceback
            print("📍 Трейс виклику:")
            traceback.print_stack()
            return {}

        if not isinstance(text, str):
            print(f"❌ КРИТИЧНА ПОМИЛКА: Параметр text не є рядком: {type(text)}")
            return {}

        if not text.strip():
            print("❌ КРИТИЧНА ПОМИЛКА: Параметр text є порожнім рядком!")
            return {}

        print(f"✅ Текст коректний, довжина: {len(text)} символів")

        # КРОК 2: Показуємо початок тексту
        preview = text[:200] + "..." if len(text) > 200 else text
        print(f"📄 Початок тексту: {repr(preview)}")

        # КРОК 3: Перевіряємо налаштування людей
        print(f"👥 Обрані люди: {self.selected_people}")
        print(f"🎭 Спеціальні ролі: {self.special_roles_selection}")

        # КРОК 4: Шукаємо плейсхолдери в тексті
        pattern = r'\{\{[^}]+\}\}'
        all_placeholders = re.findall(pattern, text)
        print(f"🔍 Знайдено плейсхолдерів в тексті: {len(all_placeholders)}")
        if all_placeholders:
            print(f"📋 Список плейсхолдерів: {all_placeholders}")

        # КРОК 5: Генеруємо замінники
        print("\n🔧 Генерація індивідуальних замінників...")
        individual_replacements = self.generate_individual_replacements(text)

        print("\n🔧 Генерація групових замінників...")
        parts_replacements = self.generate_parts_replacements(text)

        # КРОК 6: Об'єднуємо результати
        all_replacements = {}
        all_replacements.update(individual_replacements)
        all_replacements.update(parts_replacements)

        print(f"\n📊 РЕЗУЛЬТАТ ГЕНЕРАЦІЇ:")
        print(f"   Індивідуальних замінників: {len(individual_replacements)}")
        print(f"   Групових замінників: {len(parts_replacements)}")
        print(f"   Загальна кількість: {len(all_replacements)}")

        # КРОК 7: Показуємо непорожні замінники
        non_empty = {k: v for k, v in all_replacements.items() if v.strip()}
        print(f"   Непорожніх замінників: {len(non_empty)}")

        if non_empty:
            print("\n📝 Активні замінники:")
            for placeholder, replacement in non_empty.items():
                short_replacement = replacement[:50] + "..." if len(replacement) > 50 else replacement
                print(f"   {placeholder} -> '{short_replacement}'")
        else:
            print("\n⚠️ ПОПЕРЕДЖЕННЯ: Немає активних замінників!")

        print("=" * 50)
        print("🏁 КІНЕЦЬ ГЕНЕРАЦІЇ ЗАМІННИКІВ")
        print("=" * 50)

        return all_replacements

    def process_document_text(self, text):
        """Обробляє текст документа: виконує заміни з повною діагностикою"""
        print("\n" + "=" * 60)
        print("📄 ПОЧАТОК ОБРОБКИ ДОКУМЕНТА")
        print("=" * 60)

        if not text:
            print("❌ КРИТИЧНА ПОМИЛКА: Текст не переданий до process_document_text!")
            return ""

        print(f"📊 Вхідний текст, довжина: {len(text)} символів")

        # Генеруємо замінники з повною діагностикою
        replacements = self.generate_replacements(text)

        if not replacements:
            print("⚠️ ПОПЕРЕДЖЕННЯ: Замінників не згенеровано! Повертаємо оригінальний текст.")
            return text

        # Виконуємо заміни
        print(f"\n🔄 Виконання замін...")
        original_text = text
        replacement_count = 0

        for placeholder, replacement in replacements.items():
            if placeholder in text:
                old_text = text
                text = text.replace(placeholder, replacement)
                if old_text != text:
                    replacement_count += 1
                    print(f"   ✅ Замінено {placeholder}")
                else:
                    print(f"   ❌ Помилка заміни {placeholder}")

        print(f"📊 Виконано замін: {replacement_count} з {len(replacements)}")

        # Очищаємо форматування
        print("🧹 Очищення форматування...")
        text = self.clean_text_formatting(text)

        print(f"📊 Фінальний результат, довжина: {len(text)} символів")
        print("=" * 60)
        print("🏁 КІНЕЦЬ ОБРОБКИ ДОКУМЕНТА")
        print("=" * 60)

        return text

    def clean_text_formatting(self, text):
        """Очищає форматування тексту після видалення плейсхолдерів"""
        # Видаляємо подвійні коми
        text = re.sub(r',\s*,', ',', text)

        # Видаляємо коми на початку рядка
        text = re.sub(r'^\s*,\s*', '', text, flags=re.MULTILINE)

        # Видаляємо коми в кінці рядка
        text = re.sub(r',\s*$', '', text, flags=re.MULTILINE)

        # Видаляємо зайві пробіли
        text = re.sub(r'[ \t]+', ' ', text)

        # Видаляємо пробіли на початку та в кінці
        text = text.strip()

        return text

    # ДІАГНОСТИЧНІ МЕТОДИ
    def diagnose_replacement_issue(self, text=None):
        """Діагностує проблеми з генерацією замінників"""
        print("🔍 === ДІАГНОСТИКА ПРОБЛЕМИ ===")

        # Перевіряємо вхідний текст
        if not text:
            print("❌ Текст не переданий або порожній!")
            print("💡 Перевірте виклик методу generate_replacements()")
            return

        print(f"✅ Текст переданий, довжина: {len(text)} символів")

        # Перевіряємо обраних людей
        print(f"👥 Обрані люди: {self.selected_people}")
        print(f"📊 Кількість обраних: {len(self.selected_people)}")

        # Перевіряємо спеціальні ролі
        print(f"🎭 Спеціальні ролі: {self.special_roles_selection}")

        # Перевіряємо плейсхолдери в тексті
        pattern = r'\{\{[^}]+\}\}'
        placeholders = re.findall(pattern, text)
        print(f"🔍 Знайдені плейсхолдери в тексті: {placeholders}")

        # Тестуємо генерацію
        if placeholders:
            individual_replacements = self.generate_individual_replacements(text)
            parts_replacements = self.generate_parts_replacements(text)

            print(f"📊 Індивідуальних замінників: {len(individual_replacements)}")
            print(f"📊 Частинних замінників: {len(parts_replacements)}")

            # Показуємо непорожні замінники
            all_replacements = {**individual_replacements, **parts_replacements}
            non_empty = {k: v for k, v in all_replacements.items() if v.strip()}

            print(f"✅ Непорожніх замінників: {len(non_empty)}")
            for placeholder, replacement in non_empty.items():
                short_replacement = replacement[:50] + "..." if len(replacement) > 50 else replacement
                print(f"   {placeholder} -> {short_replacement}")

        print("🔍 === КІНЕЦЬ ДІАГНОСТИКИ ===")

    def test_with_sample_text(self):
        """Тестує роботу з зразковим текстом"""
        print("🧪 === ТЕСТ З ЗРАЗКОВИМ ТЕКСТОМ ===")

        sample_text = """
        Тестовий документ:
        {{PERSON_BASAI}}
        {{PERSON_MOKINA}}
        {{SELECTED_PEOPLE_PART_1}}
        {{SELECTED_PEOPLE_PART_2}}
        {{MATERIAL_RESPONSIBLE}}
        {{SELECTED_PEOPLE_LIST}}
        """

        print("🔍 Тестуємо діагностику:")
        self.diagnose_replacement_issue(sample_text)

        print("\n📄 Тестуємо обробку:")
        result = self.process_document_text(sample_text)

        print(f"\n📋 Результат:\n{result}")

        return result

    # Решта методів залишається без змін...
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

    # Методи для сумісності
    def add_person(self, person_data):
        pass

    def update_person(self, index, person_data):
        pass

    def delete_person(self, index):
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
        """Повертає дані людини за індексом"""
        people_list = self.get_people()
        if 0 <= index < len(people_list):
            return people_list[index]
        return None

    def get_person_count(self):
        """Повертає кількість людей"""
        return len(PEOPLE_CONFIG)

    def __del__(self):
        """Деструктор для очистки after() задач"""
        try:
            self.cleanup_after_jobs()
        except Exception:
            pass


# Глобальний екземпляр менеджера людей
people_manager = PeopleManager()

# Тестування
if __name__ == "__main__":
    pm = PeopleManager()
    pm.selected_people.add("basai")
    pm.selected_people.add("mokina")

    # Запускаємо новий тест
    print("🧪 === ТЕСТУВАННЯ НОВИХ ФУНКЦІЙ ===")
    pm.test_with_sample_text()