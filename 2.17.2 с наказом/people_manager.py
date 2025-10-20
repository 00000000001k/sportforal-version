# people_manager.py

import json
import os
from utils import get_executable_dir

# Конфігурація людей
PEOPLE_CONFIG = {
    "mokina": {
        "name": "Ольга МОКІНА",
        "position": "Головний спеціаліст\r\nфізкультурно-оздоровчої роботи\r\nсеред ветеранів війни",
        "display_name": "Ольга МОКІНА (Головний спеціаліст)",
        "placeholders": {
            "full_block": "{{PERSON_MOKINA}}",
            "name_only": "{{MOKINA_NAME}}",
            "position_only": "{{MOKINA_POSITION}}"
        }
    },
    "bulavko": {
        "name": "Костянтин БУЛАВКО",
        "position": "Провідний спеціаліст\r\nфізкультурно-оздоровчої роботи\r\nсеред ветеранів війни",
        "display_name": "Костянтин БУЛАВКО (Провідний спеціаліст)",
        "placeholders": {
            "full_block": "{{PERSON_BULAVKO}}",
            "name_only": "{{BULAVKO_NAME}}",
            "position_only": "{{BULAVKO_POSITION}}"
        }
    },
    "basai": {
        "name": "Микола БАСАЙ",
        "position": "Завідуючий сектором\r\nфізкультурно-оздоровчої роботи\r\nсеред ветеранів війни",
        "display_name": "Микола БАСАЙ (Завідуючий сектором)",
        "placeholders": {
            "full_block": "{{PERSON_BASAI}}",
            "name_only": "{{BASAI_NAME}}",
            "position_only": "{{BASAI_POSITION}}"
        }
    },
    "gordeeva": {
        "name": "Вікторія ГОРДЄЄВА",
        "position": "Провідний спеціаліст сектору\r\nфізкультурно-оздоровчої\r\nта спортивно-масової роботи",
        "display_name": "Вікторія ГОРДЄЄВА (Провідний спеціаліст)",
        "placeholders": {
            "full_block": "{{PERSON_GORDEEVA}}",
            "name_only": "{{GORDEEVA_NAME}}",
            "position_only": "{{GORDEEVA_POSITION}}"
        }
    }
}

# Спеціальні ролі
SPECIAL_ROLES = {
    "material_responsible": {
        "title": "Матеріально-відповідальна особа",
        "placeholder": "{{MATERIAL_RESPONSIBLE}}",
        "options": ["basai", "gordeeva"],
        "default": "basai"
    }
}


class PeopleManager:
    def __init__(self):
        self.settings_file = os.path.join(get_executable_dir(), "people_settings.json")
        self.selected_people = set()
        self.special_roles_selection = {}
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

    def generate_replacements(self):
        """Генерує словник замін для шаблонів"""
        replacements = {}

        # Спочатку додаємо всі плейсхолдери для видалення
        # Для необраних людей - видаляємо весь абзац/рядок з плейсхолдером
        for person_id, person_data in PEOPLE_CONFIG.items():
            if person_id not in self.selected_people:
                # Видаляємо весь рядок з плейсхолдером
                replacements[person_data['placeholders']['full_block']] = ""
                replacements[person_data['placeholders']['name_only']] = ""
                replacements[person_data['placeholders']['position_only']] = ""

        # Додаємо спеціальні ролі як порожні (якщо не використовуються)
        for role_config in SPECIAL_ROLES.values():
            replacements[role_config['placeholder']] = ""

        # Тепер обробляємо тільки обраних людей
        for person_id in self.selected_people:
            person_data = PEOPLE_CONFIG.get(person_id)
            if person_data:
                # Повний блок з правильним форматуванням Word
                # Використовуємо \r\n для переносів і точні пробіли для відступів
                full_block = (
                    f"{person_data['position']}\r\n"
                    f"\r\n"
                    f"“___”________________2025 року 				{person_data['name']}\r\n"
                    f""
                )
                replacements[person_data['placeholders']['full_block']] = full_block

                # Окремо ім'я та посада
                replacements[person_data['placeholders']['name_only']] = person_data['name']
                replacements[person_data['placeholders']['position_only']] = person_data['position']

        # Обробляємо спеціальні ролі
        for role_id, role_config in SPECIAL_ROLES.items():
            selected_person_id = self.get_special_role(role_id)
            person_data = PEOPLE_CONFIG.get(selected_person_id)

            if person_data:
                # Для матеріально-відповідальної особи
                if role_id == "material_responsible":
                    material_block = (
                        f"{person_data['position']}\r\n"
                        f"\r\n"
                        f"“___”________________2025 року 				{person_data['name']}"
                    )
                    replacements[role_config['placeholder']] = material_block

        return replacements

    def get_replacements_for_removal(self):
        """Генерує словник для видалення необраних людей з документа"""
        replacements = {}

        # Знаходимо всіх необраних людей і замінюємо їх плейсхолдери на порожні рядки
        for person_id, person_data in PEOPLE_CONFIG.items():
            if person_id not in self.selected_people:
                # Видаляємо весь блок цієї людини
                replacements[person_data['placeholders']['full_block']] = ""
                replacements[person_data['placeholders']['name_only']] = ""
                replacements[person_data['placeholders']['position_only']] = ""

        return replacements

    def get_selected_count(self):
        """Повертає кількість обраних людей"""
        return len(self.selected_people)

    def get_summary(self):
        """Повертає короткий опис обраних людей"""
        if not self.selected_people:
            return "Жодна особа не обрана"

        names = []
        for person_id in self.selected_people:
            person_data = PEOPLE_CONFIG.get(person_id)
            if person_data:
                names.append(person_data['name'])

        material_person_id = self.get_special_role("material_responsible")
        material_person = PEOPLE_CONFIG.get(material_person_id)
        material_name = material_person['name'] if material_person else "Не вибрано"

        summary = f"Обрано осіб: {len(names)}\n"
        summary += f"Матвідповідальна: {material_name}"

        return summary

    # Нові методи для сумісності з новою системою
    def add_person(self, person_data):
        """Додає нову людину (заглушка для сумісності)"""
        # Цей метод може бути розширений для динамічного додавання людей
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
                'id': person_id
            })
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


# Глобальний екземпляр менеджера людей
people_manager = PeopleManager()