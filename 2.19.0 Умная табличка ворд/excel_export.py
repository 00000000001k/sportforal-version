# excel_export.py
"""
Основной модуль для экспорта данных в Excel
"""

import re
import traceback
import tkinter.messagebox as messagebox

# Импорты из наших модулей
from excel_config import get_headers_config, get_ordered_headers, get_headers_mapping, is_numeric_field, \
    convert_to_number
from excel_data_processor import (
    ensure_file_structure, create_new_file, add_new_row, update_existing_row,
    get_product_name_from_row_data
)

# Попытка импорта улучшенной логики
try:
    from excel_update_logic import process_data_row_improved

    IMPROVED_LOGIC_AVAILABLE = True
    print("[INFO] Улучшенная логика обновления загружена")
except ImportError:
    IMPROVED_LOGIC_AVAILABLE = False
    print("[WARNING] Модуль улучшенной логики не найден, используется базовая логика")

    # Импортируем базовую логику из нашего модуля
    from excel_data_processor import process_data_row as process_data_row_basic

# Попытка импорта обработчика ошибок
try:
    from error_handler import log_and_show_error
except ImportError:
    def log_and_show_error(exc_type, exc_value, exc_traceback, error_log_file="error.txt"):
        print(f"Error (logging stub): {exc_value}")
        traceback.print_exception(exc_type, exc_value, exc_traceback)
        try:
            messagebox.showerror("Помилка (Excel Export)", f"Сталася помилка: {str(exc_value)}")
        except:
            pass


# ----------------------------------------------------------------------------------------------------------------------
# Основная функция экспорта
def export_document_data_to_excel(document_blocks,
                                  fields_list,
                                  output_filename="база_даних.xlsx",
                                  tabview=None,
                                  update_mode=True,
                                  similarity_threshold=70):
    """
    Экспорт данных в Excel с поддержкой обновления существующего файла

    Args:
        document_blocks: список блоков документов
        fields_list: список полей (теперь не используется, заменен на конфигурацию)
        output_filename: имя выходного файла
        tabview: объект tabview для получения номеров событий
        update_mode: если True - добавляет к существующему файлу, если False - перезаписывает
        similarity_threshold: порог схожести для поиска дублей (по умолчанию 70)
    """
    print(f"[DEBUG] ===== ПОЧАТОК export_document_data_to_excel =====")
    print(f"[DEBUG] Кількість блоків: {len(document_blocks)}")
    print(f"[DEBUG] Файл виводу: {output_filename}")
    print(f"[DEBUG] Режим обновления: {update_mode}")

    # Инициализация переменных статистики
    enable_statistics = True
    similarity_stats = {
        'similarities': [],
        'updates_by_similarity': {},
        'additions_by_similarity': {}
    }

    try:
        # Получаем конфигурацию заголовков
        headers = get_ordered_headers(fields_list)
        headers_mapping = get_headers_mapping(fields_list)

        print(f"[DEBUG] Заголовки: {headers}")
        print(f"[DEBUG] Поля из fields_list: {fields_list}")

        # Создаем словарь для быстрого поиска номеров событий
        event_numbers_map = {}
        if tabview:
            try:
                if hasattr(tabview, '_tab_dict'):
                    for tab_name in tabview._tab_dict.keys():
                        try:
                            tab = tabview.tab(tab_name)
                            event_number = getattr(tab, 'event_number', None)
                            event_numbers_map[tab_name] = event_number
                            print(f"[DEBUG] Событие '{tab_name}' имеет номер: {event_number}")
                        except Exception as e:
                            print(f"[DEBUG] Не удалось получить номер для '{tab_name}': {e}")
                            event_numbers_map[tab_name] = ""
                else:
                    print("[DEBUG] tabview не имеет _tab_dict")
            except Exception as e:
                print(f"[ERROR] Ошибка получения номеров событий: {e}")

        # Подготавливаем файл
        if update_mode:
            wb, ws = ensure_file_structure(output_filename, fields_list)
        else:
            # Режим перезаписи - создаем новый файл
            wb, ws = create_new_file(output_filename, fields_list)

        print(f"[DEBUG] Режим обновления: {update_mode}")

        # Обрабатываем данные
        processed_blocks = 0
        updated_records = 0
        added_records = 0

        # Статистика по причинам действий (для улучшенной логики)
        action_stats = {
            'no_event_number': 0,
            'no_match_found': 0,
            'similarity_update': 0,
            'too_many_changes': 0
        }

        for i, block in enumerate(document_blocks, 1):
            tab_name = block.get("tab_name", "")
            print(f"[DEBUG] Обробляємо блок {i}: tab_name='{tab_name}'")

            # Получаем номер события
            event_number = ""
            if 'event_number' in block:
                event_number = block['event_number']
            else:
                event_number = event_numbers_map.get(tab_name, "")

            # Подготавливаем данные для строки
            entries = block.get("entries", {})
            row_data = []

            # Заполняем данные согласно конфигурации заголовков
            for header_config in get_headers_config(fields_list):
                key = header_config["key"]

                if key == "event_number":
                    value = event_number if event_number is not None else ""
                elif key == "event_name":
                    value = tab_name
                else:
                    # Ищем значение в entries по ключу из fields_list
                    entry_widget_or_value = entries.get(key)
                    if hasattr(entry_widget_or_value, 'get'):
                        value = entry_widget_or_value.get()
                    else:
                        value = entry_widget_or_value if entry_widget_or_value is not None else ""

                # Обрабатываем числовые поля
                if is_numeric_field(key) and value != "":
                    converted_value = convert_to_number(value)
                    print(
                        f"[DEBUG] Числовое поле '{key}': '{value}' -> {converted_value} (тип: {type(converted_value)})")
                    row_data.append(converted_value)
                else:
                    row_data.append(value)
                    print(f"[DEBUG] Поле '{key}': '{value}'")

            # Получаем название товара для логики обновления
            product_name = get_product_name_from_row_data(row_data, headers)
            print(f"[DEBUG] Товар для блока: '{product_name}'")

            # Добавляем/обновляем строку в листе
            if update_mode:
                # Используем улучшенную логику
                action, row_num, reason = process_data_row_improved(
                    ws, row_data, headers, event_number, product_name,
                    add_new_row, update_existing_row, is_numeric_field,
                    similarity_threshold
                )

                # Обновляем статистику действий
                if action == 'updated':
                    updated_records += 1
                    # Детализируем причины обновления
                    if 'similarity' in reason:
                        action_stats['similarity_update'] += 1
                        similarity_match = re.search(r'similarity_(\d+\.?\d*)%', reason)
                        if similarity_match:
                            sim_score = float(similarity_match.group(1))
                            similarity_stats['similarities'].append(sim_score)
                            similarity_bucket = f"{int(sim_score // 10) * 10}-{int(sim_score // 10) * 10 + 9}%"
                            similarity_stats['updates_by_similarity'][similarity_bucket] = \
                                similarity_stats['updates_by_similarity'].get(similarity_bucket, 0) + 1

                elif action == 'added':
                    added_records += 1
                    # Детализируем причины добавления
                    if 'no_event_number' in reason:
                        action_stats['no_event_number'] += 1
                    elif 'no_match_found' in reason:
                        action_stats['no_match_found'] += 1
                    elif 'significant_changes' in reason:
                        action_stats['too_many_changes'] += 1
                        similarity_match = re.search(r'similarity_(\d+\.?\d*)%', reason)
                        if similarity_match:
                            sim_score = float(similarity_match.group(1))
                            similarity_stats['similarities'].append(sim_score)
                            similarity_bucket = f"{int(sim_score // 10) * 10}-{int(sim_score // 10) * 10 + 9}%"
                            similarity_stats['additions_by_similarity'][similarity_bucket] = \
                                similarity_stats['additions_by_similarity'].get(similarity_bucket, 0) + 1

                print(f"[INFO] Блок {i}: {action.upper()} строка {row_num}, причина: {reason}")

            else:
                # Режим перезаписи - просто добавляем все строки
                row_num = add_new_row(ws, row_data)
                added_records += 1
                print(f"[INFO] Блок {i}: ДОБАВЛЕНА строка {row_num}")

            processed_blocks += 1

        # Сохраняем файл
        try:
            wb.save(output_filename)
            print(f"[SUCCESS] ✅ Файл сохранен: {output_filename}")
        except Exception as e:
            print(f"[ERROR] ❌ Ошибка сохранения файла: {e}")
            return False

        # Выводим детальную статистику
        print(f"\n[INFO] 📊 СТАТИСТИКА ОБРАБОТКИ:")
        print(f"[INFO]   Обработано блоков: {processed_blocks}")
        print(f"[INFO]   Обновлено записей: {updated_records}")
        print(f"[INFO]   Добавлено новых записей: {added_records}")
        print(f"[INFO]   Всего операций: {updated_records + added_records}")

        if enable_statistics and update_mode:
            print(f"\n[INFO] 📈 ДЕТАЛЬНАЯ СТАТИСТИКА ДЕЙСТВИЙ:")
            print(f"[INFO]   Без номера события: {action_stats['no_event_number']}")
            print(f"[INFO]   Не найдено совпадений: {action_stats['no_match_found']}")
            print(f"[INFO]   Обновления по схожести: {action_stats['similarity_update']}")
            print(f"[INFO]   Добавления из-за больших изменений: {action_stats['too_many_changes']}")

            # Статистика по схожести
            if similarity_stats['similarities']:
                avg_similarity = sum(similarity_stats['similarities']) / len(similarity_stats['similarities'])
                print(f"\n[INFO] 🎯 СТАТИСТИКА СХОЖЕСТИ:")
                print(f"[INFO]   Средняя схожесть обработанных записей: {avg_similarity:.1f}%")
                print(f"[INFO]   Количество сравнений: {len(similarity_stats['similarities'])}")

                # Распределение обновлений по схожести
                if similarity_stats['updates_by_similarity']:
                    print(f"[INFO]   Обновления по диапазонам схожести:")
                    for bucket, count in sorted(similarity_stats['updates_by_similarity'].items()):
                        print(f"[INFO]     {bucket}: {count} обновлений")

                # Распределение добавлений по схожести
                if similarity_stats['additions_by_similarity']:
                    print(f"[INFO]   Добавления по диапазонам схожести:")
                    for bucket, count in sorted(similarity_stats['additions_by_similarity'].items()):
                        print(f"[INFO]     {bucket}: {count} добавлений")

        # Рекомендации по оптимизации
        if enable_statistics and processed_blocks > 0:
            print(f"\n[INFO] 💡 РЕКОМЕНДАЦИИ:")

            # Анализ эффективности порога схожести
            if similarity_stats['similarities']:
                avg_sim = sum(similarity_stats['similarities']) / len(similarity_stats['similarities'])
                if avg_sim > similarity_threshold + 10:
                    print(
                        f"[SUGGEST] Рассмотрите увеличение порога схожести до {min(90, int(avg_sim))}% для более точного поиска дублей")
                elif avg_sim < similarity_threshold - 10:
                    print(
                        f"[SUGGEST] Рассмотрите уменьшение порога схожести до {max(60, int(avg_sim))}% для лучшего объединения записей")

            # Анализ соотношения обновлений к добавлениям
            if added_records + updated_records > 0:
                update_ratio = updated_records / (added_records + updated_records) * 100
                if update_ratio < 20:
                    print(
                        f"[SUGGEST] Низкий процент обновлений ({update_ratio:.1f}%) - возможно, стоит проверить качество данных или настроить пороги")
                elif update_ratio > 80:
                    print(
                        f"[SUGGEST] Высокий процент обновлений ({update_ratio:.1f}%) - возможно, данные сильно дублируются")

            # Проверка на проблемы с номерами событий
            if action_stats['no_event_number'] > processed_blocks * 0.1:
                print(f"[WARNING] Более 10% блоков без номеров событий - проверьте логику получения номеров")

        print(f"[DEBUG] ===== КОНЕЦ export_document_data_to_excel =====")
        return True

    except Exception as e:
        print(f"[ERROR] ❌ Критическая ошибка в export_document_data_to_excel: {e}")
        import traceback
        traceback.print_exc()
        return False