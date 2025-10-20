# sportforall/excel_logger.py

import openpyxl
import os
import datetime

# Імпортуємо власні модулі
from sportforall import constants
from sportforall import error_handling
from sportforall.models import Event, Contract

# Заголовки для стовпців у файлі Excel
EXCEL_HEADERS = [
    "Час генерації",
    "Назва Заходу",
    "ID Заходу", # Може бути корисно для ідентифікації
    "Назва Договору",
    "ID Договору", # Може бути корисно для ідентифікації
    "Контрагент", # З поля '<контрагент>' договору
    "Дата Договору", # З поля '<дата>' договору
    "Загальна Сума (грн)", # З поля '<разом>' договору
    "Шлях до згенерованого файлу",
]

def log_contract(event: Event, contract: Contract, generated_filepath: str):
    """
    Записує інформацію про згенерований договір у файл Excel журналу.
    Якщо файл не існує, створює його з заголовками.

    Args:
        event: Об'єкт Event, до якого належить договір. Може бути None, якщо логується без заходу.
        contract: Об'єкт Contract, який був згенерований. Очікується об'єкт Contract.
        generated_filepath: Повний шлях до згенерованого файлу документа.
    """
    # Перевіряємо, чи отримали об'єкт договору
    if not isinstance(contract, Contract):
        error_logger.log_error(TypeError("Передано не об'єкт Contract до log_contract"), f"Помилка логування: Очікувався об'єкт Contract, отримано {type(contract)}. Логування пропущено.")
        print("Помилка логування: Некоректний об'єкт договору. Перевірте error.txt")
        return

    log_filepath = constants.EXCEL_LOG_FILE

    try:
        # Перевіряємо, чи файл журналу існує
        if not os.path.exists(log_filepath):
            # Якщо файл не існує, створюємо нову книгу та аркуш
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Журнал Генерації"
            # Записуємо заголовки
            sheet.append(EXCEL_HEADERS)
            print(f"Створено новий файл журналу Excel: {log_filepath}")
        else:
            # Якщо файл існує, відкриваємо його
            # read_only=False (за замовчуванням), keep_vba=False (за замовчуванням)
            workbook = openpyxl.load_workbook(log_filepath)
            sheet = workbook.active # Обираємо активний аркуш (перший)

        # Підготовка даних для нового рядка
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Отримуємо дані з об'єктів, використовуючи перевірки None та .get() для словника fields
        event_name = event.name if event and isinstance(event, Event) else "Невідомий Захід"
        event_id = event.id if event and isinstance(event, Event) else ""
        contract_name = contract.name # Назва договору завжди має бути
        contract_id = contract.id # ID договору завжди має бути

        # Отримуємо значення з полів договору, використовуючи ключі-плейсхолдери
        # Використовуємо .get() з порожнім рядком за замовчуванням на випадок, якщо поле відсутнє
        counterparty = contract.fields.get("<контрагент>", "")
        contract_date = contract.fields.get("<дата>", "")
        total_sum = contract.fields.get(constants.TOTAL_SUM_PLACEHOLDER, "") # Беремо розраховану суму з полів

        # Дані для нового рядка у тому ж порядку, що й заголовки
        new_row_data = [
            timestamp,
            event_name,
            event_id,
            contract_name,
            contract_id,
            counterparty,
            contract_date,
            total_sum,
            generated_filepath,
        ]

        # Додаємо новий рядок даних до аркуша
        sheet.append(new_row_data)

        # Зберігаємо книгу
        workbook.save(log_filepath)
        print(f"Інформацію про договір '{contract_name}' записано у журнал Excel '{os.path.basename(log_filepath)}'.")

    except Exception as e:
        # Логуємо помилку у файл логів додатка
        error_logger.log_error(e, f"Помилка при записі у файл журналу Excel: {log_filepath}")
        print(f"Помилка: Не вдалося записати інформацію у файл журналу Excel. Перевірте {constants.ERROR_FILE}")


# Приклад використання (для тестування)
if __name__ == "__main__":
    # Створюємо фіктивні об'єкти Event та Contract для тестування
    class MockEvent:
        def __init__(self, name, id):
            self.name = name
            self.id = id
        # Додаємо заглушки to_dict(), щоб не було помилок, якщо вони викликаються
        def to_dict(self): return {"name": self.name, "id": self.id}


    class MockContract:
        def __init__(self, name, id, fields):
            self.name = name
            self.id = id
            self.fields = fields
        # Додаємо заглушки to_dict(), щоб не було помилок, якщо вони викликаються
        def to_dict(self): return {"name": self.name, "id": self.id, "fields": self.fields}


    test_event = MockEvent("Тестовий Захід Логування", "event_123")
    test_contract_1 = MockContract(
        "Тестовий Договір 1",
        "contract_abc",
        {
            "<контрагент>": "ТОВ \"Тест\"",
            "<дата>": "2025-05-16",
            constants.TOTAL_SUM_PLACEHOLDER: "15000,50", # Використовуємо кому як в GUI для тесту
            "<інше поле>": "Яке не логується",
        }
    )
    test_contract_2 = MockContract(
        "Тестовий Договір 2 (Без контрагента)",
        "contract_def",
        {
            "<дата>": "2025-05-17",
            constants.TOTAL_SUM_PLACEHOLDER: "500,00", # Використовуємо кому як в GUI для тесту
        }
    )
    # Тестовий договір без будь-яких заповнених полів
    test_contract_3 = MockContract(
        "Тестовий Договір 3 (Порожній)",
        "contract_xyz",
        {} # Порожній словник полів
    )


    test_filepath_1 = "temp_generated_docs/Договір Тестовий Договір 1.docx"
    test_filepath_2 = "temp_generated_docs/Договір Тестовий Договір 2.docx"
    test_filepath_3 = "temp_generated_docs/Договір Тестовий Договір 3.docx"


    # Визначаємо шлях до тестового файлу журналу
    test_log_file = "test_" + constants.EXCEL_LOG_FILE

    print(f"Тестуємо логування у файл: {test_log_file}")

    # Змінюємо константу EXCEL_LOG_FILE на час тестування, щоб не перезаписати реальний журнал
    # Це небезпечно в реальному коді, але для тесту - прийнятно.
    # Краще передавати шлях до лог-файлу як параметр, якщо це можливо.
    original_log_file_constant = constants.EXCEL_LOG_FILE
    constants.EXCEL_LOG_FILE = test_log_file

    # Видаляємо тестовий файл перед кожним тестом, щоб почати з чистого аркуша
    if os.path.exists(test_log_file):
         os.remove(test_log_file)
         print(f"Видалено існуючий тестовий файл журналу: {test_log_file}")

    try:
        # Логуємо перший договір
        log_contract(test_event, test_contract_1, test_filepath_1)

        # Логуємо другий договір
        log_contract(test_event, test_contract_2, test_filepath_2)

        # Логуємо третій (порожній) договір
        log_contract(test_event, test_contract_3, test_filepath_3)

        # Логуємо договір без об'єкта заходу (тест на None)
        log_contract(None, test_contract_1, test_filepath_1 + ".2")

        # Тест з некоректним об'єктом договору
        log_contract(test_event, "Некоректний об'єкт", "fake_path.docx")


        print(f"\nТестування логування завершено. Перевірте файл '{test_log_file}'.")

    except Exception as e:
        print(f"Виникла помилка під час тестового запуску excel_logger: {e}")
        # error_logger.log_error(e, "Помилка під час тестового запуску excel_logger") # Логування вже відбувається всередині log_contract

    finally:
         # Відновлюємо оригінальне значення константи
         constants.EXCEL_LOG_FILE = original_log_file_constant
         # Очистка тестового файлу (опціонально, залиште для перевірки)
         # if os.path.exists(test_log_file):
         #      os.remove(test_log_file)
         #      print(f"Тестовий файл '{test_log_file}' видалено.")
         pass # Залиште файл для перевірки після тесту