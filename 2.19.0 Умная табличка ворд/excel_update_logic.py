# excel_update_logic.py
"""
Модуль для улучшенной логики обновления Excel файлов
Содержит функции для интеллектуального определения необходимости
обновления существующих записей или добавления новых
Использует современную библиотеку RapidFuzz вместо устаревшей fuzzywuzzy
"""

import difflib
import re

# Попытка импорта RapidFuzz, если не установлена - используем fallback
try:
    from rapidfuzz import fuzz
    RAPIDFUZZ_AVAILABLE = True
    print("[INFO] Используется RapidFuzz для продвинутого сравнения текстов")
except ImportError:
    RAPIDFUZZ_AVAILABLE = False
    print("[WARNING] RapidFuzz не установлена, используется базовое сравнение")
    print("[INFO] Для установки выполните: pip install rapidfuzz")


def normalize_text(text):
    """
    Нормализует текст для сравнения:
    - Убирает лишние пробелы
    - Приводит к нижнему регистру
    - Убирает знаки препинания
    - Обрабатывает украинские символы
    """
    if not text or text == "":
        return ""

    # Приводим к строке
    text = str(text).strip().lower()

    # Убираем множественные пробелы
    text = re.sub(r'\s+', ' ', text)

    # Убираем некоторые знаки препинания в конце/начале
    text = text.strip('.,!?;:-()[]{}\"\'')

    # Убираем дополнительные символы которые могут мешать сравнению
    text = re.sub(r'[^\w\s\u0400-\u04FF]', '', text)  # Оставляем только буквы, цифры, пробелы и кириллицу

    return text


def calculate_similarity_score(text1, text2):
    """
    Вычисляет коэффициент схожести между двумя текстами
    Возвращает число от 0 до 100 (100 = полное совпадение)
    Использует RapidFuzz для более точного и быстрого сравнения
    """
    if not text1 and not text2:
        return 100
    if not text1 or not text2:
        return 0

    # Нормализуем тексты
    norm_text1 = normalize_text(text1)
    norm_text2 = normalize_text(text2)

    # Если после нормализации тексты одинаковые
    if norm_text1 == norm_text2:
        return 100

    if RAPIDFUZZ_AVAILABLE:
        # Используем RapidFuzz для более точного сравнения
        # Коэффициент соотношения (учитывает порядок слов)
        ratio_score = fuzz.ratio(norm_text1, norm_text2)

        # Частичное соотношение (один текст содержится в другом)
        partial_score = fuzz.partial_ratio(norm_text1, norm_text2)

        # Сортированное соотношение (игнорирует порядок слов)
        token_sort_score = fuzz.token_sort_ratio(norm_text1, norm_text2)

        # Соотношение набора токенов (более продвинутый алгоритм)
        token_set_score = fuzz.token_set_ratio(norm_text1, norm_text2)

        # Берем максимальный результат из всех методов
        max_score = max(ratio_score, partial_score, token_sort_score, token_set_score)

        print(f"[DEBUG] Сравнение '{text1}' и '{text2}':")
        print(f"[DEBUG]   Нормализованные: '{norm_text1}' и '{norm_text2}'")
        print(f"[DEBUG]   Коэффициенты: ratio={ratio_score:.1f}, partial={partial_score:.1f}")
        print(f"[DEBUG]   token_sort={token_sort_score:.1f}, token_set={token_set_score:.1f}")
        print(f"[DEBUG]   Итоговый коэффициент: {max_score:.1f}")

        return max_score
    else:
        # Fallback на стандартную библиотеку
        similarity = difflib.SequenceMatcher(None, norm_text1, norm_text2).ratio() * 100
        print(f"[DEBUG] Базовое сравнение '{text1}' и '{text2}': {similarity:.1f}%")
        return similarity


def find_best_matching_row(ws, event_number, product_name, headers, similarity_threshold=80):
    """
    Находит наиболее подходящую строку для обновления:
    1. Сначала ищет точное совпадение по номеру события и товару
    2. Если не найдено, ищет строки с тем же номером события и похожим товаром
    3. Возвращает строку с наивысшим коэффициентом схожести (если > threshold)

    Args:
        ws: рабочий лист Excel
        event_number: номер события
        product_name: название товара
        headers: заголовки таблицы
        similarity_threshold: минимальный порог схожести (по умолчанию 80%)

    Returns:
        tuple: (номер_строки, коэффициент_схожести) или (None, 0) если не найдено
    """
    if not event_number or event_number == "":
        return None, 0

    # Проверяем, что есть данные
    if ws.max_row < 2:
        return None, 0

    # Находим индексы нужных колонок
    event_number_col = None
    product_col = None

    for col_idx, header in enumerate(headers, 1):
        if header.lower() in ["номер заходу", "номер события", "event_number"]:
            event_number_col = col_idx
        elif header.lower() in ["товар", "товару", "product", "продукт"]:
            product_col = col_idx

    if event_number_col is None:
        print(f"[ERROR] Не найдена колонка номера события")
        return None, 0

    best_row = None
    best_similarity = 0

    # Ищем все строки с совпадающим номером события
    matching_rows = []
    for row in range(2, ws.max_row + 1):
        event_cell_value = ws.cell(row=row, column=event_number_col).value
        if str(event_cell_value).strip() == str(event_number).strip():
            matching_rows.append(row)

    if not matching_rows:
        print(f"[DEBUG] Не найдено строк с номером события '{event_number}'")
        return None, 0

    print(f"[DEBUG] Найдено {len(matching_rows)} строк с номером события '{event_number}'")

    # Если нет колонки товара или товар не указан
    if product_col is None or not product_name:
        print(f"[DEBUG] Нет колонки товара или товар не указан, возвращаем первую найденную строку")
        return matching_rows[0], 100

    # Сравниваем товары в найденных строках
    for row in matching_rows:
        existing_product = ws.cell(row=row, column=product_col).value
        if existing_product is None:
            existing_product = ""

        similarity = calculate_similarity_score(existing_product, product_name)

        print(f"[DEBUG] Строка {row}: товар '{existing_product}', схожесть {similarity:.1f}%")

        if similarity > best_similarity:
            best_similarity = similarity
            best_row = row

    # Проверяем порог схожести
    if best_similarity >= similarity_threshold:
        print(f"[INFO] Найдена подходящая строка {best_row} со схожестью {best_similarity:.1f}%")
        return best_row, best_similarity
    else:
        print(f"[INFO] Максимальная схожесть {best_similarity:.1f}% меньше порога {similarity_threshold}%")
        return None, 0


def analyze_field_changes(old_value, new_value, field_name, is_numeric_field_func):
    """
    Анализирует изменения в конкретном поле и определяет их значимость

    Args:
        old_value: старое значение
        new_value: новое значение
        field_name: название поля
        is_numeric_field_func: функция для проверки числовых полей

    Returns:
        dict: информация об изменении
    """
    if old_value is None:
        old_value = ""
    if new_value is None:
        new_value = ""

    # Нормализуем значения для сравнения
    old_normalized = normalize_text(str(old_value))
    new_normalized = normalize_text(str(new_value))

    change_info = {
        'field': field_name,
        'old_value': old_value,
        'new_value': new_value,
        'is_changed': old_normalized != new_normalized,
        'is_significant': False,
        'similarity': 100,
        'change_type': 'no_change'
    }

    if not change_info['is_changed']:
        return change_info

    # Определяем тип изменения
    if is_numeric_field_func(field_name):
        # Для числовых полей любое изменение значимо
        change_info['is_significant'] = True
        change_info['change_type'] = 'numeric_change'
        change_info['similarity'] = 0 if old_normalized != new_normalized else 100
    else:
        # Для текстовых полей проверяем степень изменения
        similarity = calculate_similarity_score(old_value, new_value)
        change_info['similarity'] = similarity

        if similarity < 70:  # Если схожесть меньше 70%, считаем изменение значимым
            change_info['is_significant'] = True
            change_info['change_type'] = 'significant_text_change'
        elif similarity < 90:
            change_info['is_significant'] = True
            change_info['change_type'] = 'moderate_text_change'
        else:
            change_info['is_significant'] = False
            change_info['change_type'] = 'minor_text_change'

    return change_info


def should_update_or_add(existing_data, new_data, headers, event_number, product_name, is_numeric_field_func):
    """
    Определяет, нужно ли обновить существующую запись или создать новую
    на основе детального анализа изменений в данных

    Args:
        existing_data: данные из существующей строки
        new_data: новые данные
        headers: заголовки таблицы
        event_number: номер события
        product_name: название товара
        is_numeric_field_func: функция для проверки числовых полей

    Returns:
        tuple: (действие, детали_анализа)
    """

    # Подсчитываем количество значимых изменений
    significant_changes = 0
    moderate_changes = 0
    minor_changes = 0
    total_fields = 0

    # Подробная информация об изменениях
    changes_details = []

    # Поля, которые мы не учитываем при принятии решения
    ignore_fields = {'event_number', 'event_name', 'номер заходу', 'название события'}

    for idx, header in enumerate(headers):
        if header.lower() in [h.lower() for h in ignore_fields]:
            continue

        if idx < len(existing_data) and idx < len(new_data):
            old_value = existing_data[idx]
            new_value = new_data[idx]

            change_info = analyze_field_changes(old_value, new_value, header, is_numeric_field_func)
            changes_details.append(change_info)

            total_fields += 1

            if change_info['is_changed']:
                if change_info['change_type'] == 'significant_text_change' or change_info['change_type'] == 'numeric_change':
                    significant_changes += 1
                elif change_info['change_type'] == 'moderate_text_change':
                    moderate_changes += 1
                else:
                    minor_changes += 1

                print(f"[DEBUG] Изменение в '{header}': {old_value} -> {new_value}")
                print(f"[DEBUG]   Тип: {change_info['change_type']}, схожесть: {change_info['similarity']:.1f}%")

    # Процент различных типов изменений
    significant_percentage = (significant_changes / total_fields * 100) if total_fields > 0 else 0
    moderate_percentage = (moderate_changes / total_fields * 100) if total_fields > 0 else 0
    total_changes_percentage = ((significant_changes + moderate_changes + minor_changes) / total_fields * 100) if total_fields > 0 else 0

    print(f"[DEBUG] Анализ изменений:")
    print(f"[DEBUG]   Значимые: {significant_changes}/{total_fields} ({significant_percentage:.1f}%)")
    print(f"[DEBUG]   Умеренные: {moderate_changes}/{total_fields} ({moderate_percentage:.1f}%)")
    print(f"[DEBUG]   Незначительные: {minor_changes}/{total_fields}")
    print(f"[DEBUG]   Всего изменений: {total_changes_percentage:.1f}%")

    # Принятие решения на основе анализа
    analysis = {
        'significant_changes': significant_changes,
        'moderate_changes': moderate_changes,
        'minor_changes': minor_changes,
        'total_fields': total_fields,
        'significant_percentage': significant_percentage,
        'moderate_percentage': moderate_percentage,
        'total_changes_percentage': total_changes_percentage,
        'changes_details': changes_details
    }

    # Логика принятия решения:
    # 1. Если значимых изменений больше 30% - добавляем новую запись
    # 2. Если значимых + умеренных изменений больше 60% - добавляем новую запись
    # 3. Иначе - обновляем существующую запись

    if significant_percentage > 30:
        decision = 'add'
        reason = f'too_many_significant_changes_{significant_percentage:.1f}%'
    elif (significant_percentage + moderate_percentage) > 60:
        decision = 'add'
        reason = f'too_many_total_changes_{significant_percentage + moderate_percentage:.1f}%'
    else:
        decision = 'update'
        reason = f'acceptable_changes_{significant_percentage:.1f}%_significant'

    print(f"[INFO] Решение: {'ОБНОВИТЬ' if decision == 'update' else 'ДОБАВИТЬ НОВУЮ ЗАПИСЬ'}")
    print(f"[INFO] Причина: {reason}")

    return decision, analysis


def get_row_data(ws, row_number, num_columns):
    """
    Получает данные из указанной строки листа Excel
    """
    row_data = []
    for col in range(1, num_columns + 1):
        value = ws.cell(row=row_number, column=col).value
        row_data.append(value)
    return row_data


def process_data_row_improved(ws, row_data, headers, event_number, product_name,
                              add_new_row_func, update_existing_row_func, is_numeric_field_func,
                              similarity_threshold=80):
    """
    Улучшенная обработка строки данных с интеллектуальным определением
    необходимости обновления или добавления новой записи
    Использует RapidFuzz для более точного сравнения

    Args:
        ws: рабочий лист Excel
        row_data: данные для записи
        headers: заголовки таблицы
        event_number: номер события
        product_name: название товара
        add_new_row_func: функция для добавления новой строки
        update_existing_row_func: функция для обновления существующей строки
        is_numeric_field_func: функция для проверки числовых полей
        similarity_threshold: минимальный порог схожести для поиска

    Returns:
        tuple: (действие, номер_строки, дополнительная_информация)
    """

    if not event_number or str(event_number).strip() == "":
        # Нет номера события - добавляем как новую запись
        print(f"[WARNING] Нет номера события, добавляем как новую запись")
        new_row = add_new_row_func(ws, row_data)
        return 'added', new_row, 'no_event_number'

    # Ищем наиболее подходящую строку для обновления
    best_row, similarity = find_best_matching_row(ws, event_number, product_name, headers, similarity_threshold)

    if best_row is None:
        # Не найдено подходящих записей - добавляем новую
        print(f"[INFO] Не найдено подходящих записей для события '{event_number}', добавляем новую")
        new_row = add_new_row_func(ws, row_data)
        return 'added', new_row, 'no_match_found'

    # Получаем существующие данные для анализа
    existing_data = get_row_data(ws, best_row, len(headers))

    # Определяем, обновлять или добавлять
    decision, analysis = should_update_or_add(existing_data, row_data, headers, event_number, product_name,
                                              is_numeric_field_func)

    if decision == 'update':
        # Обновляем существующую запись
        update_existing_row_func(ws, best_row, row_data, headers)
        reason = f"similarity_{similarity:.1f}%_changes_{analysis['significant_percentage']:.1f}%"
        return 'updated', best_row, reason
    else:
        # Добавляем новую запись
        new_row = add_new_row_func(ws, row_data)
        reason = f"similarity_{similarity:.1f}%_significant_changes_{analysis['significant_percentage']:.1f}%"
        return 'added', new_row, reason


def get_similarity_statistics(ws, headers, similarity_threshold=80):
    """
    Анализирует существующие данные и возвращает статистику схожести
    Полезно для настройки порогов схожести

    Args:
        ws: рабочий лист Excel
        headers: заголовки таблицы
        similarity_threshold: порог схожести для анализа

    Returns:
        dict: статистика схожести
    """
    if ws.max_row < 3:  # Нет данных для анализа
        return {'status': 'insufficient_data'}

    # Находим колонку товара
    product_col = None
    for col_idx, header in enumerate(headers, 1):
        if header.lower() in ["товар", "товару", "product", "продукт"]:
            product_col = col_idx
            break

    if product_col is None:
        return {'status': 'no_product_column'}

    # Собираем все названия товаров
    products = []
    for row in range(2, ws.max_row + 1):
        product = ws.cell(row=row, column=product_col).value
        if product:
            products.append(str(product))

    if len(products) < 2:
        return {'status': 'insufficient_products'}

    # Анализируем схожесть между товарами
    similarities = []
    for i in range(len(products)):
        for j in range(i + 1, len(products)):
            similarity = calculate_similarity_score(products[i], products[j])
            similarities.append({
                'product1': products[i],
                'product2': products[j],
                'similarity': similarity
            })

    # Статистика
    similarity_values = [s['similarity'] for s in similarities]
    high_similarity = [s for s in similarities if s['similarity'] >= similarity_threshold]

    stats = {
        'status': 'success',
        'total_products': len(products),
        'total_comparisons': len(similarities),
        'high_similarity_pairs': len(high_similarity),
        'avg_similarity': sum(similarity_values) / len(similarity_values),
        'max_similarity': max(similarity_values),
        'min_similarity': min(similarity_values),
        'threshold': similarity_threshold,
        'high_similarity_examples': high_similarity[:5]  # Первые 5 примеров
    }

    return stats