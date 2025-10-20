# excel_config.py
"""
Конфигурация заголовков, столбцов и порядка полей для Excel экспорта
"""

def get_column_order():
    """
    Определяет правильный порядок столбцов
    Можно легко изменять последовательность здесь
    """
    column_order = [
        "event_number",  # Номер заходу
        "event_name",  # Захід
        "дата",  # Дата
        "адреса",  # Адреса
        "дк",  # ДК
        "товар",  # Товар
        "кількість",  # Кількість
        "ціна за одиницю",  # Ціна за одиницю
        "разом"  # Разом
    ]
    return column_order


def get_numeric_fields():
    """
    Возвращает список полей, которые должны быть обработаны как числовые
    """
    numeric_fields = [
        "кількість",
        "ціна за одиницю",
        "разом",
        "количество",
        "цена за единицу",
        "итого",
        "сумма",
        "цена",
        "стоимость"
    ]
    return [field.lower() for field in numeric_fields]


def is_numeric_field(field_name):
    """
    Проверяет, является ли поле числовым
    """
    numeric_fields = get_numeric_fields()
    return field_name.lower() in numeric_fields


def convert_to_number(value):
    """
    Конвертирует значение в число, если это возможно
    Возвращает число или исходное значение, если конвертация невозможна
    """
    if value is None or value == "":
        return 0

    # Если уже число
    if isinstance(value, (int, float)):
        return value

    # Если строка
    if isinstance(value, str):
        # Убираем пробелы
        value = value.strip()

        # Пустая строка = 0
        if not value:
            return 0

        # Убираем апострофы в начале (если есть)
        if value.startswith("'"):
            value = value[1:]

        # Заменяем запятые на точки для десятичных чисел
        value = value.replace(",", ".")

        try:
            # Пытаемся конвертировать в float
            num_value = float(value)

            # Если это целое число, возвращаем как int
            if num_value.is_integer():
                return int(num_value)
            else:
                return num_value

        except (ValueError, AttributeError):
            # Если не удалось конвертировать, возвращаем исходное значение
            print(f"[WARNING] Не удалось конвертировать '{value}' в число")
            return value

    return value


def sort_fields_by_priority(fields_list):
    """
    Сортирует поля согласно заданному порядку
    Поля, которых нет в приоритетном списке, добавляются в конец
    """
    column_order = get_column_order()

    # Создаем словарь для быстрого поиска приоритета
    # Исключаем базовые поля event_number и event_name из приоритетного списка
    priority_map = {}
    for idx, col in enumerate(column_order):
        col_lower = col.lower()
        if col_lower not in ["event_number", "event_name"]:
            priority_map[col_lower] = idx

    # Разделяем поля на приоритетные и обычные
    priority_fields = []
    other_fields = []

    for field in fields_list:
        field_lower = field.lower()
        if field_lower in priority_map:
            priority_fields.append((priority_map[field_lower], field))
        else:
            other_fields.append(field)

    # Сортируем приоритетные поля по порядку
    priority_fields.sort(key=lambda x: x[0])
    sorted_priority = [field for _, field in priority_fields]

    # Сортируем остальные поля по алфавиту
    other_fields.sort()

    # Объединяем
    sorted_fields = sorted_priority + other_fields

    print(f"[DEBUG] Исходные поля: {fields_list}")
    print(f"[DEBUG] Отсортированные поля: {sorted_fields}")

    return sorted_fields


def get_base_headers_config():
    """
    Возвращает базовые заголовки (номер события и название)
    """
    base_headers = [
        {"key": "event_number", "title": "Номер заходу", "width": 15},
        {"key": "event_name", "title": "Захід", "width": 25}
    ]
    return base_headers


def get_field_width(field_name):
    """
    Определяет оптимальную ширину колонки в зависимости от названия поля
    """
    field_lower = field_name.lower()

    # Специальные ширины для конкретных полей
    width_map = {
        "дата": 12,
        "адреса": 30,
        "дк": 20,
        "товар": 25,
        "кількість": 12,
        "ціна за одиницю": 18,
        "разом": 15,
        "номер": 15,
        "телефон": 15,
        "email": 25
    }

    # Ищем совпадения в названии поля
    for key, width in width_map.items():
        if key in field_lower:
            return width

    # По умолчанию
    return 20


def get_headers_config(fields_list):
    """
    Возвращает полную конфигурацию заголовков:
    - Базовые заголовки (номер события, название)
    - Отсортированные динамические заголовки из fields_list (без дублирования)
    """
    # Базовые заголовки
    headers_config = get_base_headers_config()

    # Получаем ключи базовых заголовков для избежания дублирования
    base_keys = {header["key"].lower() for header in headers_config}
    base_titles = {header["title"].lower() for header in headers_config}

    print(f"[DEBUG] Базовые ключи: {base_keys}")
    print(f"[DEBUG] Базовые заголовки: {base_titles}")

    # Фильтруем fields_list, исключая поля, которые уже есть в базовых заголовках
    filtered_fields = []
    for field in fields_list:
        field_lower = field.lower()

        # Проверяем, не является ли это поле базовым заголовком
        if (field_lower not in base_keys and
                field_lower not in base_titles and
                field_lower not in ["event_number", "event_name", "номер заходу", "захід"]):
            filtered_fields.append(field)
        else:
            print(f"[DEBUG] Исключаем дублирующееся поле: '{field}'")

    print(f"[DEBUG] Отфильтрованные поля: {filtered_fields}")

    # Сортируем отфильтрованные поля по приоритету
    sorted_fields = sort_fields_by_priority(filtered_fields)

    # Добавляем отсортированные поля из filtered_fields
    for field in sorted_fields:
        # Определяем ширину колонки в зависимости от типа поля
        width = get_field_width(field)

        headers_config.append({
            "key": field,
            "title": field,
            "width": width
        })

    print(f"[DEBUG] Итоговая конфигурация заголовков:")
    for i, header in enumerate(headers_config):
        print(f"[DEBUG]   {i + 1}. {header['key']} -> '{header['title']}' (ширина: {header['width']})")

    return headers_config


def get_ordered_headers(fields_list):
    """Возвращает только заголовки в правильном порядке"""
    return [header["title"] for header in get_headers_config(fields_list)]


def get_headers_mapping(fields_list):
    """Возвращает маппинг ключ -> заголовок"""
    return {header["key"]: header["title"] for header in get_headers_config(fields_list)}


def get_column_widths(fields_list):
    """Возвращает ширины колонок"""
    return {header["title"]: header["width"] for header in get_headers_config(fields_list)}