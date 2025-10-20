# excel_config.py
"""
Конфигурация заголовков, столбцов и порядка полей для Excel экспорта
"""

from fields_config import get_fields_config, get_numeric_fields


def get_exportable_fields():
    """
    Возвращает только те поля, которые нужно экспортировать в Excel
    Отсортированные по порядку (order)
    """
    fields_config = get_fields_config()

    # Фильтруем только экспортируемые поля
    exportable = []
    for key, config in fields_config.items():
        if config.get("export", False):
            exportable.append({
                "key": key,
                "title": config["title"],
                "width": config["width"],
                "order": config["order"]
            })

    # Сортируем по порядку
    exportable.sort(key=lambda x: x["order"])

    print(f"[DEBUG] Экспортируемые поля в порядке:")
    for i, field in enumerate(exportable, 1):
        print(f"[DEBUG]   {i}. {field['key']} -> '{field['title']}' (ширина: {field['width']})")

    return exportable


def get_excluded_fields():
    """
    Возвращает список полей, которые исключены из экспорта
    """
    fields_config = get_fields_config()

    excluded = []
    for key, config in fields_config.items():
        if not config.get("export", False):
            excluded.append(key)

    print(f"[DEBUG] Исключенные из экспорта поля: {excluded}")
    return excluded


def is_field_exportable(field_name):
    """
    Проверяет, нужно ли экспортировать это поле
    """
    fields_config = get_fields_config()
    field_lower = field_name.lower()

    # Ищем точное совпадение
    if field_lower in fields_config:
        return fields_config[field_lower].get("export", False)

    # Если поля нет в конфигурации, не экспортируем его
    print(f"[WARNING] Поле '{field_name}' не найдено в конфигурации, исключаем из экспорта")
    return False


def filter_exportable_fields(fields_list):
    """
    Фильтрует список полей, оставляя только те, которые нужно экспортировать
    """
    filtered_fields = []

    for field in fields_list:
        if is_field_exportable(field):
            filtered_fields.append(field)
        else:
            print(f"[DEBUG] Исключаем поле из экспорта: '{field}'")

    print(f"[DEBUG] Исходные поля: {fields_list}")
    print(f"[DEBUG] Отфильтрованные поля: {filtered_fields}")

    return filtered_fields


def sort_fields_by_config(fields_list):
    """
    Сортирует поля согласно конфигурации (по параметру order)
    """
    fields_config = get_fields_config()

    # Создаем список с порядковыми номерами
    fields_with_order = []

    for field in fields_list:
        field_lower = field.lower()
        if field_lower in fields_config:
            order = fields_config[field_lower].get("order", 999)
            fields_with_order.append((order, field))
        else:
            # Неизвестные поля в конец
            fields_with_order.append((999, field))

    # Сортируем по order
    fields_with_order.sort(key=lambda x: x[0])
    sorted_fields = [field for _, field in fields_with_order]

    print(f"[DEBUG] Исходные поля: {fields_list}")
    print(f"[DEBUG] Отсортированные поля: {sorted_fields}")

    return sorted_fields


def get_base_headers():
    """
    Возвращает базовые заголовки, которые всегда должны быть в экспорте
    (независимо от того, есть ли они в данных)
    """
    fields_config = get_fields_config()
    base_headers = []

    # Добавляем базовые поля
    base_fields = ["event_number", "event_name"]

    for field in base_fields:
        if field in fields_config and fields_config[field].get("export", False):
            config = fields_config[field]
            base_headers.append({
                "key": field,
                "title": config["title"],
                "width": config["width"],
                "order": config["order"]
            })

    return base_headers


def get_headers_config(fields_list):
    """
    Возвращает конфигурацию заголовков для экспорта
    Включает базовые заголовки + поля из данных
    """
    fields_config = get_fields_config()

    # Получаем базовые заголовки (всегда включаются)
    base_headers = get_base_headers()
    base_keys = [header["key"].lower() for header in base_headers]

    # Фильтруем поля из данных (исключаем базовые, чтобы не дублировать)
    data_fields = [field for field in fields_list if field.lower() not in base_keys]

    # Фильтруем только экспортируемые поля из данных
    exportable_data_fields = filter_exportable_fields(data_fields)

    # Формируем заголовки для полей из данных
    data_headers = []
    for field in exportable_data_fields:
        field_lower = field.lower()

        if field_lower in fields_config:
            config = fields_config[field_lower]
            data_headers.append({
                "key": field,
                "title": config["title"],
                "width": config["width"],
                "order": config["order"]
            })
        else:
            # Для неизвестных полей используем дефолтные значения
            data_headers.append({
                "key": field,
                "title": field,
                "width": 20,
                "order": 999
            })

    # Объединяем базовые и обычные заголовки
    all_headers = base_headers + data_headers

    # Сортируем по order
    all_headers.sort(key=lambda x: x["order"])

    print(f"[DEBUG] Базовые заголовки: {[h['key'] for h in base_headers]}")
    print(f"[DEBUG] Поля из данных (после фильтрации): {[h['key'] for h in data_headers]}")
    print(f"[DEBUG] Итоговая конфигурация заголовков:")
    for i, header in enumerate(all_headers, 1):
        print(f"[DEBUG]   {i}. {header['key']} -> '{header['title']}' (ширина: {header['width']})")

    return all_headers


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


def get_row_data_with_base_fields(row_data, event_number, event_name):
    """
    Добавляет базовые поля к данным строки
    """
    # Создаем копию данных строки
    enhanced_row = row_data.copy() if row_data else {}

    # Добавляем базовые поля
    enhanced_row["event_number"] = event_number
    enhanced_row["event_name"] = event_name

    return enhanced_row


def format_row_for_export(row_data, headers_config, event_number=None, event_name=None):
    """
    Форматирует строку данных для экспорта согласно конфигурации заголовков
    """
    # Добавляем базовые поля, если они переданы
    if event_number is not None or event_name is not None:
        row_data = get_row_data_with_base_fields(row_data, event_number, event_name)

    formatted_row = []

    for header in headers_config:
        field_key = header["key"]
        value = row_data.get(field_key, "")

        # Конвертируем числовые поля
        if is_numeric_field(field_key):
            value = convert_to_number(value)

        formatted_row.append(value)

    return formatted_row


def get_headers_titles(fields_list):
    """
    Возвращает список заголовков для экспорта
    """
    return [header["title"] for header in get_headers_config(fields_list)]


def get_ordered_headers(fields_list):
    """
    Возвращает список заголовков для экспорта (алиас для get_headers_titles)
    Функция для обратной совместимости
    """
    return get_headers_titles(fields_list)


def get_headers_mapping(fields_list):
    """Возвращает маппинг ключ -> заголовок"""
    return {header["key"]: header["title"] for header in get_headers_config(fields_list)}


def get_column_widths(fields_list):
    """Возвращает ширины колонок"""
    return {header["title"]: header["width"] for header in get_headers_config(fields_list)}