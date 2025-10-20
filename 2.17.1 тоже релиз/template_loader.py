# template_loader.py

import os
import sys

def get_executable_dir():
    """Повертає папку де знаходиться exe файл або скрипт"""
    if getattr(sys, 'frozen', False):
        # Якщо запущено як exe файл (PyInstaller, cx_Freeze, etc.)
        return os.path.dirname(sys.executable)
    else:
        # Якщо запущено як Python скрипт
        return os.path.dirname(os.path.abspath(__file__))

# Папка templates поруч з exe файлом
TEMPLATE_FOLDER = os.path.join(get_executable_dir(), "templates")


def get_available_templates():
    """Повертає словник шаблонів: {ім'я: шлях} з автооновленням"""
    templates = {}

    # Перевіряємо чи існує папка templates
    if not os.path.exists(TEMPLATE_FOLDER):
        try:
            os.makedirs(TEMPLATE_FOLDER)
            print(f"[INFO] Створена папка шаблонів: {TEMPLATE_FOLDER}")
        except Exception as e:
            print(f"[ERROR] Не вдалося створити папку шаблонів: {e}")
            return {}

    try:
        # Отримуємо всі файли з папки
        files = os.listdir(TEMPLATE_FOLDER)

        for filename in files:
            # Перевіряємо розширення файлу
            if filename.lower().endswith(('.docx', '.docm')):
                # Отримуємо ім'я без розширення
                name = os.path.splitext(filename)[0]
                path = os.path.join(TEMPLATE_FOLDER, filename)

                # Перевіряємо чи файл існує та доступний для читання
                if os.path.isfile(path) and os.access(path, os.R_OK):
                    templates[name] = path
                else:
                    print(f"[WARNING] Файл {filename} недоступний для читання")

        print(f"[INFO] Знайдено {len(templates)} валідних шаблонів: {list(templates.keys())}")

    except Exception as e:
        print(f"[ERROR] Помилка при читанні папки шаблонів: {e}")
        return {}

    return templates


def get_template_path(template_name):
    """Повертає шлях до конкретного шаблону"""
    templates = get_available_templates()
    return templates.get(template_name)


def is_template_valid(template_path):
    """Перевіряє чи валідний шаблон"""
    if not template_path or not os.path.exists(template_path):
        return False

    if not template_path.lower().endswith(('.docx', '.docm')):
        return False

    return os.access(template_path, os.R_OK)


def get_templates_folder():
    """Повертає шлях до папки з шаблонами"""
    return TEMPLATE_FOLDER


def refresh_templates_info():
    """Повертає детальну інформацію про шаблони"""
    templates = get_available_templates()

    info = {
        'folder_path': TEMPLATE_FOLDER,
        'folder_exists': os.path.exists(TEMPLATE_FOLDER),
        'executable_dir': get_executable_dir(),
        'total_templates': len(templates),
        'templates': templates
    }

    return info


def open_templates_folder():
    """Відкриває папку з шаблонами в провіднику файлів"""
    try:
        if os.name == 'nt':  # Windows
            os.startfile(TEMPLATE_FOLDER)
        elif os.name == 'posix':  # macOS і Linux
            os.system(f'open "{TEMPLATE_FOLDER}"')  # macOS
            # або os.system(f'xdg-open "{TEMPLATE_FOLDER}"')  # Linux
    except Exception as e:
        print(f"[ERROR] Не вдалося відкрити папку: {e}")


# Додаткова функція для демонстрації
def print_folder_info():
    """Виводить інформацію про папки для діагностики"""
    print(f"Executable directory: {get_executable_dir()}")
    print(f"Templates folder: {TEMPLATE_FOLDER}")
    print(f"Templates folder exists: {os.path.exists(TEMPLATE_FOLDER)}")
    print(f"Current working directory: {os.getcwd()}")
    print(f"sys.frozen: {getattr(sys, 'frozen', False)}")
    print(f"sys.executable: {sys.executable}")