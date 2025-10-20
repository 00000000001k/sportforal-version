
import json
import os
import sys
import traceback
from pathlib import Path

# Налаштування шляхів
MEMORY_FILE = "contracts_memory.json"
STATE_FILE = Path("data/state.json")

def make_json_serializable(obj):
    if isinstance(obj, dict):
        return {k: make_json_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [make_json_serializable(i) for i in obj]
    elif isinstance(obj, (str, int, float, bool)) or obj is None:
        return obj
    elif hasattr(obj, "get") and callable(obj.get):
        try:
            return obj.get()
        except Exception:
            return f"[ErrorInGet:{type(obj).__name__}]"
    elif callable(obj):
        return f"[Callable:{type(obj).__name__}]"
    else:
        return str(obj)


# Обробка помилок
try:
    from error_handler import log_error_to_file, log_and_show_error
except ImportError:
    # Заглушка, якщо error_handler не знайдено
    def log_error_to_file(exc_type, exc_value, exc_traceback, error_log_file="error.txt"):
        try:
            # Обертаємо в Exception, якщо exc_value не є винятком
            if not isinstance(exc_value, BaseException):
                exc_value = Exception(str(exc_value))
            with open(error_log_file, "a", encoding="utf-8") as f:
                traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)
        except Exception as log_exc:
            print("[log_error_to_file] Ошибка при записи в лог:", log_exc)
            traceback.print_exception(exc_type, exc_value, exc_traceback)

    def log_and_show_error(exc_type, exc_value, exc_traceback):
        if not isinstance(exc_value, BaseException):
            exc_value = Exception(str(exc_value))
        print(f"Error: {exc_value}")
        traceback.print_exception(exc_type, exc_value, exc_traceback)


# Основні функції для роботи з пам'яттю договорів
def load_memory(memory_file_path=MEMORY_FILE):
    try:
        if os.path.exists(memory_file_path):
            with open(memory_file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}
    except Exception as e:
        log_error_to_file(type(e), e, sys.exc_info()[2])
        return {}


def save_memory(data_to_save, template_path=None, memory_file_path=MEMORY_FILE):
    try:
        current_data = load_memory(memory_file_path)
        if template_path:
            norm_path = os.path.normpath(template_path)
            current_data.setdefault(norm_path, {})
            if isinstance(current_data[norm_path], dict):
                current_data[norm_path].update(data_to_save)
            else:
                current_data[norm_path] = data_to_save
        else:
            current_data.update(data_to_save)

        with open(memory_file_path, "w", encoding="utf-8") as f:
            json.dump(current_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_error_to_file(type(e), e, sys.exc_info()[2])



def get_template_memory(template_path, memory_file_path=MEMORY_FILE):
    try:
        memory_data = load_memory(memory_file_path)
        norm_path = os.path.normpath(template_path)
        return memory_data.get(norm_path, {})
    except Exception as e:
        log_error_to_file(type(e), e, sys.exc_info()[2])
        return {}


def save_full_state(app_state):
    try:
        os.makedirs("data", exist_ok=True)
        # Гарантовано серіалізуємо, навіть якщо хтось передав неочищений словник
        serializable_state = make_json_serializable(app_state)
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(serializable_state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_and_show_error(type(e), e, sys.exc_info()[2])



def load_full_state():
    try:
        if STATE_FILE.exists():
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return {}
                return json.loads(content)
    except Exception as e:
        log_and_show_error(type(e), e, sys.exc_info()[2])
    return {}

def save_current_state(tabview, document_blocks):
    """Зберігає поточний стан: назви вкладок та введені поля."""
    try:
        data = {
            "tabs": tabview._name_list if hasattr(tabview, "_name_list") else [],
            "document_blocks": []
        }

        for block in document_blocks:
            block_data = {
                "path": block.get("path"),
                "entries": {key: entry.get() for key, entry in block.get("entries", {}).items()}
            }
            data["document_blocks"].append(block_data)

        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    except Exception as e:
        print(f"Помилка при збереженні стану: {e}")

def load_saved_state():
    """Завантажує стан вкладок та полів, якщо такий є."""
    if not os.path.exists(STATE_FILE):
        return None

    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Помилка при завантаженні стану: {e}")
        return None
