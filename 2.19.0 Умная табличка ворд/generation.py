# generation.py
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import tkinter.messagebox as messagebox
from tkinter import filedialog
import pythoncom
import win32com.client as win32
import re

from globals import document_blocks
from data_persistence import save_memory
from excel_export import export_document_data_to_excel
from text_utils import number_to_ukrainian_text
import koshtorys
from error_handler import log_and_show_error
from people_manager import people_manager
from koshtorys import fill_koshtorys


def enhanced_extract_placeholders_from_word(template_path):
    """
    Витягує всі плейсхолдери типу <поле> з документу Word з детальною діагностикою
    Додано підтримку плейсхолдера <таблиця_товарів>
    """
    placeholders = set()
    word_app = None
    doc = None

    try:
        pythoncom.CoInitialize()
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False

        template_path_abs = os.path.abspath(template_path)
        if not os.path.exists(template_path_abs):
            print(f"[ERROR] Шаблон не знайдено: {template_path_abs}")
            return placeholders

        doc = word_app.Documents.Open(template_path_abs)

        # Отримуємо весь текст з документу
        full_text = doc.Content.Text
        # print(f"[DEBUG] Документ {template_path}: довжина тексту {len(full_text)} символів")

        # Також перевіряємо таблиці окремо
        table_text = ""
        for table in doc.Tables:
            for row in table.Rows:
                for cell in row.Cells:
                    cell_content = cell.Range.Text
                    table_text += " " + cell_content
                    full_text += " " + cell_content

        # if table_text:
        #     print(f"[DEBUG] Знайдено текст у таблицях: {len(table_text)} символів")

        # Шукаємо всі плейсхолдери типу <поле>
        pattern = r'<([^>]+)>'
        matches = re.findall(pattern, full_text)
        # print(f"[DEBUG] Знайдено raw matches: {matches}")

        for match in matches:
            # Очищуємо від спеціальних символів Word
            clean_match = match.strip().replace('\r', '').replace('\x07', '').replace('\n', '')
            if clean_match:
                placeholders.add(clean_match)
                # print(f"[DEBUG] Додано плейсхолдер: '{clean_match}'")

        # Додаткова перевірка на випадок, якщо плейсхолдери мають нестандартний формат
        # Шукаємо варіанти з фігурними дужками
        pattern2 = r'\{([^}]+)\}'
        matches2 = re.findall(pattern2, full_text)
        for match in matches2:
            clean_match = match.strip().replace('\r', '').replace('\x07', '').replace('\n', '')
            if clean_match and any(
                    keyword in clean_match.lower() for keyword in ['people', 'person', 'selected', 'part']):
                placeholders.add(f"{{{clean_match}}}")
                # print(f"[DEBUG] Додано плейсхолдер з фігурними дужками: '{{{clean_match}}}'")

        # print(f"[DEBUG] Фінальні плейсхолдери в {template_path}: {sorted(placeholders)}")

    except Exception as e:
        print(f"[ERROR] Помилка при витягуванні плейсхолдерів з {template_path}: {e}")
        traceback.print_exc()
    finally:
        if doc:
            try:
                doc.Close(False)
            except:
                pass
        if word_app:
            try:
                word_app.Quit()
            except:
                pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    return placeholders


def get_all_placeholders_from_blocks(blocks):
    """
    Отримує всі унікальні плейсхолдери з усіх блоків документів
    """
    all_placeholders = set()

    for block in blocks:
        if "path" in block and block["path"] and os.path.exists(block["path"]):
            block_placeholders = enhanced_extract_placeholders_from_word(block["path"])
            all_placeholders.update(block_placeholders)

            # Зберігаємо плейсхолдери для кожного блоку
            block["placeholders"] = block_placeholders

    return sorted(list(all_placeholders))


def validate_items_data(items_data):
    """
    Перевірка даних товарів
    """
    if not items_data or len(items_data) == 0:
        raise ValueError("Будь ласка, заповніть поле для товару чи товарів")

    # Перевіряємо, що всі товари мають необхідні поля
    for i, item in enumerate(items_data, 1):
        if not item.get("товар", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено найменування")
        if not item.get("дк", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено код ДК-021:2015")
        if not item.get("кількість", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено кількість")
        if not item.get("ціна", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено ціну за одиницю")

    return True


def create_products_table(doc, items_data):
    """
    Створює таблицю товарів для вставки в Word документ
    Структура таблиці відповідно до технічного завдання
    """
    try:
        # Створюємо таблицю: заголовок + товари + підсумковий рядок
        rows_count = 2 + len(items_data)  # заголовок + товари + підсумок
        cols_count = 6  # № п/п, Найменування, ДК-021:2015, Одиниця виміру+Кількість, Ціна за одиницю, Загальна сума

        table = doc.Tables.Add(doc.Range(), rows_count, cols_count)

        # Налаштовуємо стиль таблиці
        table.Borders.Enable = True
        table.PreferredWidthType = win32.constants.wdPreferredWidthPercent
        table.PreferredWidth = 100

        # ЗАГОЛОВОК ТАБЛИЦІ (перший рядок)
        header_row = table.Rows(1)
        header_row.Cells(1).Range.Text = "№\nп/п"
        header_row.Cells(2).Range.Text = "Найменування"
        header_row.Cells(3).Range.Text = "ДК-021:2015"
        header_row.Cells(4).Range.Text = "Одиниця виміру"
        header_row.Cells(5).Range.Text = "Кількість"
        header_row.Cells(6).Range.Text = "Ціна за\nодиницю, грн."
        header_row.Cells(7).Range.Text = "Загальна\nсума, грн."

        # Додаємо ще одну колонку для правильної структури
        table.Columns.Add()

        # Перерозподіляємо заголовки після додавання колонки
        header_row.Cells(1).Range.Text = "№\nп/п"
        header_row.Cells(2).Range.Text = "Найменування"
        header_row.Cells(3).Range.Text = "ДК-021:2015"
        header_row.Cells(4).Range.Text = "Одиниця виміру"
        header_row.Cells(5).Range.Text = "Кількість"
        header_row.Cells(6).Range.Text = "Ціна за\nодиницю, грн."
        header_row.Cells(7).Range.Text = "Загальна\nсума, грн."

        # Об'єднуємо колонки "Одиниця виміру" і "Кількість" у заголовку
        # За технічним завданням ці колонки мають бути БЕЗ вертикальних ліній
        try:
            header_row.Cells(4).Merge(header_row.Cells(5))
            header_row.Cells(4).Range.Text = "Одиниця виміру │ Кількість"
        except:
            # Якщо об'єднання не вдалося, залишаємо як є
            pass

        # Форматування заголовку
        header_row.Range.Font.Bold = True
        header_row.Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

        # ЗАПОВНЕННЯ РЯДКІВ ТОВАРІВ
        total_sum = 0.0

        for i, item in enumerate(items_data):
            row_num = i + 2  # +2 тому що перший рядок - заголовок
            row = table.Rows(row_num)

            # № п/п
            row.Cells(1).Range.Text = str(i + 1)
            row.Cells(1).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

            # Найменування
            row.Cells(2).Range.Text = item.get("товар", "")

            # ДК-021:2015
            row.Cells(3).Range.Text = item.get("дк", "")
            row.Cells(3).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

            # Одиниця виміру + Кількість (об'єднана колонка)
            unit_quantity_text = f"шт. │ {item.get('кількість', '')}"
            row.Cells(4).Range.Text = unit_quantity_text
            row.Cells(4).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

            # Ціна за одиницю
            price_text = item.get("ціна", "0")
            row.Cells(5).Range.Text = price_text
            row.Cells(5).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight

            # Загальна сума (розрахунок)
            try:
                qty_str = item.get("кількість", "0").replace(",", ".")
                price_str = item.get("ціна", "0").replace(",", ".")
                qty = float(qty_str) if qty_str else 0
                price = float(price_str) if price_str else 0
                item_sum = qty * price
                total_sum += item_sum

                row.Cells(6).Range.Text = f"{item_sum:.2f}"
                row.Cells(6).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
            except (ValueError, TypeError):
                row.Cells(6).Range.Text = "0.00"
                row.Cells(6).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight

        # ПІДСУМКОВИЙ РЯДОК "РАЗОМ"
        total_row_num = len(items_data) + 2
        total_row = table.Rows(total_row_num)

        # Об'єднуємо перші 5 колонок для тексту "Разом:"
        try:
            total_row.Cells(1).Merge(total_row.Cells(5))
            total_row.Cells(1).Range.Text = "Разом:"
            total_row.Cells(1).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
            total_row.Cells(1).Range.Font.Bold = True

            # Загальна сума
            total_row.Cells(2).Range.Text = f"{total_sum:.2f}"
            total_row.Cells(2).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
            total_row.Cells(2).Range.Font.Bold = True
        except:
            # Якщо об'єднання не вдалося, заповнюємо останню колонку
            total_row.Cells(5).Range.Text = "Разом:"
            total_row.Cells(5).Range.Font.Bold = True
            total_row.Cells(6).Range.Text = f"{total_sum:.2f}"
            total_row.Cells(6).Range.Font.Bold = True

        print(f"[DEBUG] Створено таблицю товарів: {len(items_data)} товарів, загальна сума: {total_sum:.2f}")
        return table

    except Exception as e:
        print(f"[ERROR] Помилка при створенні таблиці товарів: {e}")
        traceback.print_exc()
        return None


def replace_table_placeholders(doc, table_placeholder, items_data):
    """
    Знаходить плейсхолдери <таблиця_товарів> і вставляє таблиці товарів
    """
    try:
        print(f"[DEBUG] Пошук плейсхолдера: {table_placeholder}")

        # Шукаємо всі входження плейсхолдера
        replacements_made = 0

        # Ітеруємо по всіх абзацах документу
        for paragraph in doc.Paragraphs:
            paragraph_text = paragraph.Range.Text.strip()

            if table_placeholder in paragraph_text:
                print(f"[DEBUG] Знайдено плейсхолдер в абзаці: {paragraph_text[:100]}")

                try:
                    # Створюємо таблицю товарів
                    products_table = create_products_table(doc, items_data)

                    if products_table:
                        # Встановлюємо позицію для вставки (замість абзацу з плейсхолдером)
                        insert_range = paragraph.Range

                        # Видаляємо текст плейсхолдера
                        insert_range.Text = ""

                        # Переміщуємо таблицю на позицію плейсхолдера
                        products_table.Range.Cut()
                        insert_range.Paste()

                        replacements_made += 1
                        print(f"[DEBUG] Успішно замінено плейсхолдер #{replacements_made}")

                        # Прериваємо після першої заміни, щоб уникнути помилок з ітерацією
                        break

                except Exception as e:
                    print(f"[ERROR] Помилка при заміні плейсхолдера: {e}")
                    continue

        # Якщо не знайшли в абзацах, шукаємо через Find
        if replacements_made == 0:
            print(f"[DEBUG] Плейсхолдер не знайдено в абзацах, використовуємо Find")

            find_obj = doc.Content.Find
            find_obj.ClearFormatting()

            while find_obj.Execute(FindText=table_placeholder):
                try:
                    print(f"[DEBUG] Знайдено плейсхолдер через Find")

                    # Зберігаємо позицію
                    selection_range = find_obj.Parent

                    # Створюємо таблицю
                    products_table = create_products_table(doc, items_data)

                    if products_table:
                        # Видаляємо плейсхолдер
                        selection_range.Text = ""

                        # Вставляємо таблицю
                        products_table.Range.Cut()
                        selection_range.Paste()

                        replacements_made += 1
                        print(f"[DEBUG] Успішно замінено плейсхолдер через Find #{replacements_made}")
                        break

                except Exception as e:
                    print(f"[ERROR] Помилка при заміні через Find: {e}")
                    break

        print(f"[DEBUG] Загалом замінено плейсхолдерів: {replacements_made}")
        return replacements_made > 0

    except Exception as e:
        print(f"[ERROR] Загальна помилка при заміні плейсхолдерів таблиць: {e}")
        traceback.print_exc()
        return False


def process_people_placeholders_in_document(doc):
    """
    Обробляє плейсхолдери людей у документі Word з детальною діагностикою
    """
    try:
        # Отримуємо замінники для людей
        people_replacements = people_manager.generate_replacements()

        if not people_replacements:
            # print("[DEBUG] Немає замінників для людей")
            return

        # print(f"[DEBUG] Обробляємо замінники людей: {people_replacements}")

        # Спочатку отримаємо весь текст документу для діагностики
        full_text = doc.Content.Text
        # print(f"[DEBUG] Перші 500 символів документу: {repr(full_text[:500])}")

        # Перевіряємо наявність кожного плейсхолдера в тексті
        for placeholder in people_replacements.keys():
            placeholder_exists = placeholder in full_text
            # print(f"[DEBUG] Плейсхолдер '{placeholder}' знайдено в тексті: {placeholder_exists}")
            if placeholder_exists:
                # Знаходимо контекст навколо плейсхолдера
                pos = full_text.find(placeholder)
                context_start = max(0, pos - 50)
                context_end = min(len(full_text), pos + len(placeholder) + 50)
                context = full_text[context_start:context_end]
                # print(f"[DEBUG] Контекст плейсхолдера '{placeholder}': {repr(context)}")

        # Обробляємо кожен плейсхолдер окремо
        for placeholder, replacement in people_replacements.items():
            try:
                # print(f"[DEBUG] Обробляємо плейсхолдер: '{placeholder}'")
                # print(f"[DEBUG] Заміна: '{replacement}' (довжина: {len(replacement)})")

                # Якщо replacement порожній - видаляємо весь абзац
                if replacement == "":
                    # print(f"[DEBUG] Видаляємо абзац з плейсхолдером: {placeholder}")

                    # Шукаємо абзац з плейсхолдером і видаляємо його повністю
                    find_obj = doc.Content.Find
                    find_obj.ClearFormatting()

                    found_count = 0
                    while find_obj.Execute(FindText=placeholder):
                        found_count += 1
                        # print(f"[DEBUG] Знайдено входження #{found_count} плейсхолдера {placeholder}")

                        try:
                            # Розширюємо виділення до всього абзацу
                            paragraph = find_obj.Parent.Paragraphs(1)
                            paragraph_text = paragraph.Range.Text
                            # print(f"[DEBUG] Видаляємо абзац: {repr(paragraph_text[:100])}")
                            paragraph.Range.Delete()
                            # print(f"[DEBUG] Абзац видалено успішно")
                            # Перериваємо цикл, оскільки абзац видалено
                            break
                        except Exception as delete_error:
                            print(f"[ERROR] Помилка при видаленні абзацу: {delete_error}")
                            break

                    if found_count == 0:
                        print(f"[WARNING] Плейсхолдер {placeholder} не знайдено для видалення")

                else:
                    # Звичайна заміна для обраних людей
                    # print(f"[DEBUG] Виконуємо заміну плейсхолдера: {placeholder}")

                    find_obj = doc.Content.Find
                    find_obj.ClearFormatting()
                    find_obj.Replacement.ClearFormatting()

                    # Спеціальна обробка для багаторядкового тексту
                    if "\r\n" in replacement:
                        word_replacement = replacement.replace("\r\n", "^p")
                        # print(f"[DEBUG] Багаторядковий текст конвертовано: {repr(word_replacement)}")
                    else:
                        word_replacement = replacement

                    # Замінюємо плейсхолдер на заміну
                    result = find_obj.Execute(
                        FindText=placeholder,
                        ReplaceWith=word_replacement,
                        Replace=win32.constants.wdReplaceAll
                    )

                    if result:
                        print(f"[DEBUG] Замінено {placeholder} -> {replacement[:50]}... (успішно)")
                    else:
                        print(f"[WARNING] Заміна {placeholder} не виконана (результат: {result})")

                        # Додаткова перевірка - можливо плейсхолдер має інший формат
                        alt_formats = [
                            placeholder.upper(),
                            placeholder.lower(),
                            placeholder.replace('_', ' '),
                            placeholder.replace('-', '_')
                        ]

                        for alt_format in alt_formats:
                            if alt_format != placeholder and alt_format in full_text:
                                # print(f"[DEBUG] Знайдено альтернативний формат: {alt_format}")
                                alt_result = find_obj.Execute(
                                    FindText=alt_format,
                                    ReplaceWith=word_replacement,
                                    Replace=win32.constants.wdReplaceAll
                                )
                                if alt_result:
                                    # print(f"[DEBUG] Альтернативна заміна успішна: {alt_format}")
                                    break

            except Exception as e:
                print(f"[ERROR] Помилка при заміні {placeholder}: {e}")
                traceback.print_exc()

        # Також обробляємо таблиці окремо
        # print(f"[DEBUG] Обробляємо таблиці, кількість: {doc.Tables.Count}")

        for table_idx, table in enumerate(doc.Tables, 1):
            # print(f"[DEBUG] Обробляємо таблицю #{table_idx}")

            for row_idx, row in enumerate(table.Rows, 1):
                for cell_idx, cell in enumerate(row.Cells, 1):
                    try:
                        cell_text = cell.Range.Text
                        original_cell_text = cell_text
                        modified = False

                        # Перевіряємо кожен плейсхолдер у тексті комірки
                        for placeholder, replacement in people_replacements.items():
                            if placeholder in cell_text:
                                # print(f"[DEBUG] Знайдено плейсхолдер {placeholder} в комірці [{row_idx},{cell_idx}]")

                                # Для таблиць також обробляємо переноси рядків
                                if replacement == "":
                                    # Видаляємо плейсхолдер з комірки
                                    cell_text = cell_text.replace(placeholder, "")
                                    # print(f"[DEBUG] Видалено плейсхолдер з комірки")
                                else:
                                    if "\r\n" in replacement:
                                        word_replacement = replacement.replace("\r\n", "\r")
                                    else:
                                        word_replacement = replacement
                                    cell_text = cell_text.replace(placeholder, word_replacement)
                                    # print(f"[DEBUG] Замінено в комірці: {placeholder} -> {word_replacement[:30]}...")

                                modified = True

                        # Якщо текст змінився, оновлюємо комірку
                        if modified:
                            cell.Range.Text = cell_text
                            # print(f"[DEBUG] Комірка [{row_idx},{cell_idx}] оновлена")
                            # print(f"[DEBUG] Було: {repr(original_cell_text[:50])}")
                            # print(f"[DEBUG] Стало: {repr(cell_text[:50])}")

                    except Exception as e:
                        print(f"[ERROR] Помилка при обробці комірки [{row_idx},{cell_idx}]: {e}")

        # print("[DEBUG] Обробка плейсхолдерів людей завершена")

    except Exception as e:
        print(f"[ERROR] Загальна помилка при обробці плейсхолдерів людей: {e}")
        traceback.print_exc()


def process_document_content(doc, block, current_fields):
    """
    Обробляє вміст документу Word - замінює плейсхолдери та обробляє людей
    ОНОВЛЕНО: додано обробку плейсхолдерів таблиць товарів
    """
    try:
        # 1. Спочатку обробляємо звичайні плейсхолдери
        block_placeholders = block.get("placeholders", current_fields)
        for key in block_placeholders:
            if key in block["entries"] and block["entries"][key]:
                find_obj = doc.Content.Find
                find_obj.ClearFormatting()
                find_obj.Replacement.ClearFormatting()
                find_obj.Execute(FindText=f"<{key}>",
                                 ReplaceWith=block["entries"][key].get(),
                                 Replace=win32.constants.wdReplaceAll)

        # 2. Потім обробляємо плейсхолдери людей
        process_people_placeholders_in_document(doc)

        # 3. НОВЕ: обробляємо плейсхолдери таблиць товарів
        items = block.get("items", [])
        if callable(items):
            items = items()

        # Перевіряємо наявність плейсхолдера таблиці товарів
        full_text = doc.Content.Text
        if "<таблиця_товарів>" in full_text:
            print(f"[DEBUG] Знайдено плейсхолдер <таблиця_товарів>, товарів: {len(items) if items else 0}")

            try:
                # Валідація даних товарів
                validate_items_data(items)

                # Заміна плейсхолдерів таблиць
                success = replace_table_placeholders(doc, "<таблиця_товарів>", items)

                if success:
                    print("[DEBUG] Таблиця товарів успішно вставлена")
                else:
                    print("[WARNING] Не вдалося вставити таблицю товарів")

            except ValueError as ve:
                # Показуємо помилку користувачу
                print(f"[ERROR] Помилка валідації товарів: {ve}")
                messagebox.showerror("Помилка даних товарів", str(ve))
                return False
            except Exception as e:
                print(f"[ERROR] Помилка при обробці таблиці товарів: {e}")
                traceback.print_exc()
                return False

        # print(f"[DEBUG] Обробка всіх플placeholder'ів завершена для документу")
        return True

    except Exception as e:
        print(f"[ERROR] Помилка при обробці плейсхолдерів: {e}")
        traceback.print_exc()
        return False


def generate_documents_word(tabview):
    """
    Генерує документи Word з підтримкою таблиць товарів
    """
    selected_event = tabview.get().strip()
    current_blocks = [block for block in document_blocks if block.get("tab_name", "").strip() == selected_event]

    print(f"[DEBUG] Вибраний захід: {selected_event}")
    print(f"[DEBUG] Доступні блоки: {[b.get('tab_name') for b in document_blocks]}")

    if not current_blocks:
        messagebox.showwarning("Увага", f"У заході '{selected_event}' немає жодного договору для генерації.")
        return False

    save_dir = filedialog.askdirectory(title="Оберіть папку для збереження документів Word")
    if not save_dir:
        return False

    # Отримуємо динамічні поля для поточних блоків
    current_fields = get_all_placeholders_from_blocks(current_blocks)
    print(f"[DEBUG] Динамічні поля: {current_fields}")

    # === ГЕНЕРАЦІЯ EXCEL НА ПОЧАТКУ ===
    try:
        excel_success = export_document_data_to_excel(current_blocks, save_dir, selected_event)
        if excel_success:
            print("[DEBUG] Excel файл успішно створено")
        else:
            print("[WARNING] Не вдалося створити Excel файл")
    except Exception as e:
        print(f"[ERROR] Помилка при створенні Excel: {e}")

    # === ОСНОВНА ГЕНЕРАЦІЯ WORD ДОКУМЕНТІВ ===
    word_app = None
    total_generated = 0
    failed_documents = []

    try:
        pythoncom.CoInitialize()
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False

        for block in current_blocks:
            try:
                print(f"\n[DEBUG] ========== Обробляємо блок: {block.get('title', 'Без назви')} ==========")

                # Перевіряємо наявність шаблону
                template_path = block.get("path", "")
                if not template_path or not os.path.exists(template_path):
                    error_msg = f"Шаблон не знайдено: {template_path}"
                    print(f"[ERROR] {error_msg}")
                    failed_documents.append(f"{block.get('title', 'Невідомий')}: {error_msg}")
                    continue

                # Перевіряємо, чи заповнені обов'язкові поля
                entries = block.get("entries", {})
                missing_fields = []

                for field in current_fields:
                    if field not in entries or not entries[field] or not entries[field].get().strip():
                        # Ігноруємо спеціальні плейсхолдери
                        if field not in ["таблиця_товарів"] and not field.startswith("{"):
                            missing_fields.append(field)

                if missing_fields:
                    error_msg = f"Не заповнені поля: {', '.join(missing_fields)}"
                    print(f"[WARNING] {error_msg}")
                    # Не зупиняємося, але логуємо попередження

                # Відкриваємо шаблон
                print(f"[DEBUG] Відкриваємо шаблон: {template_path}")
                doc = word_app.Documents.Open(template_path)

                # Обробляємо вміст документу (включно з таблицями товарів)
                process_success = process_document_content(doc, block, current_fields)

                if not process_success:
                    print(f"[ERROR] Помилка при обробці вмісту документу")
                    failed_documents.append(f"{block.get('title', 'Невідомий')}: Помилка обробки вмісту")
                    doc.Close(False)
                    continue

                # Генеруємо ім'я файлу
                safe_title = "".join(
                    c for c in block.get("title", "document") if c.isalnum() or c in (' ', '-', '_')).rstrip()
                if not safe_title:
                    safe_title = f"document_{total_generated + 1}"

                save_path = os.path.join(save_dir, f"{safe_title}.docx")

                # Перевіряємо, чи файл вже існує
                counter = 1
                original_save_path = save_path
                while os.path.exists(save_path):
                    name_part = safe_title
                    save_path = os.path.join(save_dir, f"{name_part}_{counter}.docx")
                    counter += 1

                # Зберігаємо документ
                print(f"[DEBUG] Зберігаємо документ: {save_path}")
                doc.SaveAs2(save_path)
                doc.Close()

                total_generated += 1
                print(f"[SUCCESS] Документ '{safe_title}' успішно згенеровано")

                # Обробляємо кошторис, якщо потрібно
                try:
                    if hasattr(koshtorys, 'fill_koshtorys') and callable(koshtorys.fill_koshtorys):
                        koshtorys_success = koshtorys.fill_koshtorys(block, save_path, word_app)
                        if koshtorys_success:
                            print(f"[DEBUG] Кошторис успішно оброблено для {safe_title}")
                        else:
                            print(f"[WARNING] Кошторис не оброблено для {safe_title}")
                except Exception as koshtorys_error:
                    print(f"[ERROR] Помилка при обробці кошторису: {koshtorys_error}")

            except Exception as e:
                error_msg = f"Помилка при генерації документу: {str(e)}"
                print(f"[ERROR] {error_msg}")
                traceback.print_exc()
                failed_documents.append(f"{block.get('title', 'Невідомий')}: {error_msg}")

                # Закриваємо документ, якщо він відкритий
                try:
                    if 'doc' in locals():
                        doc.Close(False)
                except:
                    pass

    except Exception as e:
        error_msg = f"Критична помилка при генерації документів: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        messagebox.showerror("Критична помилка", error_msg)
        return False

    finally:
        # Закриваємо Word
        if word_app:
            try:
                word_app.Quit()
            except:
                pass

        try:
            pythoncom.CoUninitialize()
        except:
            pass

    # === ЗВІТ ПРО РЕЗУЛЬТАТИ ===
    try:
        save_memory()
        print("[DEBUG] Дані збережено в пам'ять")
    except Exception as e:
        print(f"[ERROR] Помилка при збереженні даних: {e}")

    # Показуємо результат користувачу
    if total_generated > 0:
        success_msg = f"Успішно згенеровано {total_generated} документів у папці:\n{save_dir}"

        if failed_documents:
            failed_msg = f"\n\nНе вдалося згенерувати ({len(failed_documents)} документів):\n"
            failed_msg += "\n".join([f"• {doc}" for doc in failed_documents[:5]])
            if len(failed_documents) > 5:
                failed_msg += f"\n• ... та ще {len(failed_documents) - 5} документів"

            success_msg += failed_msg
            messagebox.showwarning("Генерація завершена з попередженнями", success_msg)
        else:
            messagebox.showinfo("Успіх", success_msg)

        return True
    else:
        error_msg = "Не вдалося згенерувати жодного документу."
        if failed_documents:
            error_msg += f"\n\nПомилки:\n" + "\n".join([f"• {doc}" for doc in failed_documents[:10]])
            if len(failed_documents) > 10:
                error_msg += f"\n• ... та ще {len(failed_documents) - 10} помилок"

        messagebox.showerror("Помилка генерації", error_msg)
        return False


def create_products_table(doc, items_data):
    """
    Створює таблицю товарів для вставки в Word документ
    Структура таблиці відповідно до технічного завдання
    """
    try:
        # Створюємо таблицю: заголовок + товари + підсумковий рядок
        rows_count = 2 + len(items_data)  # заголовок + товари + підсумок
        cols_count = 7  # № п/п, Найменування, ДК-021:2015, Одиниця виміру, Кількість, Ціна за одиницю, Загальна сума

        table = doc.Tables.Add(doc.Range(), rows_count, cols_count)

        # Налаштовуємо стиль таблиці
        table.Borders.Enable = True
        table.PreferredWidthType = win32.constants.wdPreferredWidthPercent
        table.PreferredWidth = 100

        # ЗАГОЛОВОК ТАБЛИЦІ (перший рядок)
        header_row = table.Rows(1)
        header_row.Cells(1).Range.Text = "№\nп/п"
        header_row.Cells(2).Range.Text = "Найменування"
        header_row.Cells(3).Range.Text = "ДК-021:2015"
        header_row.Cells(4).Range.Text = "Одиниця виміру"
        header_row.Cells(5).Range.Text = "Кількість"
        header_row.Cells(6).Range.Text = "Ціна за\nодиницю, грн."
        header_row.Cells(7).Range.Text = "Загальна\nсума, грн."

        # Об'єднуємо колонки "Одиниця виміру" і "Кількість" у заголовку
        try:
            header_row.Cells(4).Merge(header_row.Cells(5))
            header_row.Cells(4).Range.Text = "Одиниця виміру │ Кількість"
        except Exception as merge_error:
            print(f"[WARNING] Не вдалося об'єднати колонки заголовку: {merge_error}")

        # Форматування заголовку
        header_row.Range.Font.Bold = True
        header_row.Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

        # ЗАПОВНЕННЯ РЯДКІВ ТОВАРІВ
        total_sum = 0.0

        for i, item in enumerate(items_data):
            row_num = i + 2  # +2 тому що перший рядок - заголовок
            row = table.Rows(row_num)

            # № п/п
            row.Cells(1).Range.Text = str(i + 1)
            row.Cells(1).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

            # Найменування
            row.Cells(2).Range.Text = item.get("товар", "")

            # ДК-021:2015
            row.Cells(3).Range.Text = item.get("дк", "")
            row.Cells(3).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

            # Одиниця виміру + Кількість (об'єднана колонка)
            unit_quantity_text = f"шт. │ {item.get('кількість', '')}"
            row.Cells(4).Range.Text = unit_quantity_text
            row.Cells(4).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter

            # Ціна за одиницю
            price_text = item.get("ціна", "0")
            row.Cells(5).Range.Text = price_text
            row.Cells(5).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight

            # Загальна сума (розрахунок)
            try:
                qty_str = item.get("кількість", "0").replace(",", ".")
                price_str = item.get("ціна", "0").replace(",", ".")
                qty = float(qty_str) if qty_str else 0
                price = float(price_str) if price_str else 0
                item_sum = qty * price
                total_sum += item_sum

                row.Cells(6).Range.Text = f"{item_sum:.2f}"
                row.Cells(6).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
            except (ValueError, TypeError):
                row.Cells(6).Range.Text = "0.00"
                row.Cells(6).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight

        # ПІДСУМКОВИЙ РЯДОК "РАЗОМ"
        total_row_num = len(items_data) + 2
        total_row = table.Rows(total_row_num)

        # Об'єднуємо перші 5 колонок для тексту "Разом:"
        try:
            total_row.Cells(1).Merge(total_row.Cells(5))
            total_row.Cells(1).Range.Text = "Разом:"
            total_row.Cells(1).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
            total_row.Cells(1).Range.Font.Bold = True

            # Загальна сума
            total_row.Cells(2).Range.Text = f"{total_sum:.2f}"
            total_row.Cells(2).Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
            total_row.Cells(2).Range.Font.Bold = True
        except Exception as total_merge_error:
            print(f"[WARNING] Не вдалося об'єднати комірки підсумкового рядка: {total_merge_error}")
            # Якщо об'єднання не вдалося, заповнюємо останні колонки
            total_row.Cells(5).Range.Text = "Разом:"
            total_row.Cells(5).Range.Font.Bold = True
            total_row.Cells(6).Range.Text = f"{total_sum:.2f}"
            total_row.Cells(6).Range.Font.Bold = True

        print(f"[DEBUG] Створено таблицю товарів: {len(items_data)} товарів, загальна сума: {total_sum:.2f}")
        return table

    except Exception as e:
        print(f"[ERROR] Помилка при створенні таблиці товарів: {e}")
        traceback.print_exc()
        return None


def replace_table_placeholders(doc, table_placeholder, items_data):
    """
    Знаходить плейсхолдери <таблиця_товарів> і вставляє таблиці товарів
    """
    try:
        print(f"[DEBUG] Пошук плейсхолдера: {table_placeholder}")

        # Шукаємо всі входження плейсхолдера
        replacements_made = 0

        # Спочатку пробуємо через Find для точного пошуку
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()

        while find_obj.Execute(FindText=table_placeholder):
            try:
                print(f"[DEBUG] Знайдено плейсхолдер через Find")

                # Зберігаємо позицію та створюємо таблицю
                selection_range = find_obj.Parent

                # Створюємо таблицю товарів
                products_table = create_products_table(doc, items_data)

                if products_table:
                    # Видаляємо плейсхолдер
                    selection_range.Text = ""

                    # Переміщуємо таблицю на позицію плейсхолдера
                    products_table.Range.Cut()
                    selection_range.Paste()

                    replacements_made += 1
                    print(f"[DEBUG] Успішно замінено плейсхолдер #{replacements_made}")
                    break
                else:
                    print(f"[ERROR] Не вдалося створити таблицю товарів")
                    break

            except Exception as e:
                print(f"[ERROR] Помилка при заміні через Find: {e}")
                break

        # Якщо через Find не знайшли, пробуємо через ітерацію абзаців
        if replacements_made == 0:
            print(f"[DEBUG] Плейсхолдер не знайдено через Find, шукаємо в абзацах")

            paragraphs_to_process = []

            # Спочатку знаходимо всі абзаци з плейсхолдером
            for paragraph in doc.Paragraphs:
                paragraph_text = paragraph.Range.Text.strip()
                if table_placeholder in paragraph_text:
                    paragraphs_to_process.append(paragraph)

            # Потім обробляємо знайдені абзаци
            for paragraph in paragraphs_to_process:
                try:
                    print(f"[DEBUG] Обробляємо абзац з плейсхолдером")

                    # Створюємо таблицю товарів
                    products_table = create_products_table(doc, items_data)

                    if products_table:
                        # Встановлюємо позицію для вставки
                        insert_range = paragraph.Range

                        # Видаляємо текст плейсхолдера
                        insert_range.Text = ""

                        # Переміщуємо таблицю на позицію плейсхолдера
                        products_table.Range.Cut()
                        insert_range.Paste()

                        replacements_made += 1
                        print(f"[DEBUG] Успішно замінено плейсхолдер в абзаці #{replacements_made}")
                        break

                except Exception as e:
                    print(f"[ERROR] Помилка при заміні в абзаці: {e}")
                    continue

        print(f"[DEBUG] Загалом замінено плейсхолдерів: {replacements_made}")
        return replacements_made > 0

    except Exception as e:
        print(f"[ERROR] Загальна помилка при заміні плейсхолдерів таблиць: {e}")
        traceback.print_exc()
        return False


def validate_items_data(items_data):
    """
    Перевірка даних товарів
    """
    if not items_data or len(items_data) == 0:
        raise ValueError("Будь ласка, заповніть поле для товару чи товарів")

    # Перевіряємо, що всі товари мають необхідні поля
    for i, item in enumerate(items_data, 1):
        if not item.get("товар", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено найменування")
        if not item.get("дк", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено код ДК-021:2015")
        if not item.get("кількість", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено кількість")
        if not item.get("ціна", "").strip():
            raise ValueError(f"Товар №{i}: не заповнено ціну за одиницю")

        # Перевіряємо, що кількість та ціна - це числа
        try:
            qty_str = item.get("кількість", "0").replace(",", ".")
            float(qty_str)
        except (ValueError, TypeError):
            raise ValueError(f"Товар №{i}: кількість має бути числом")

        try:
            price_str = item.get("ціна", "0").replace(",", ".")
            float(price_str)
        except (ValueError, TypeError):
            raise ValueError(f"Товар №{i}: ціна має бути числом")

    return True