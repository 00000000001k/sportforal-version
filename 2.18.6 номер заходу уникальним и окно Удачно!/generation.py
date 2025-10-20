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
        #print(f"[DEBUG] Документ {template_path}: довжина тексту {len(full_text)} символів")

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
        #print(f"[DEBUG] Знайдено raw matches: {matches}")

        for match in matches:
            # Очищуємо від спеціальних символів Word
            clean_match = match.strip().replace('\r', '').replace('\x07', '').replace('\n', '')
            if clean_match:
                placeholders.add(clean_match)
                #print(f"[DEBUG] Додано плейсхолдер: '{clean_match}'")

        # Додаткова перевірка на випадок, якщо плейсхолдери мають нестандартний формат
        # Шукаємо варіанти з фігурними дужками
        pattern2 = r'\{([^}]+)\}'
        matches2 = re.findall(pattern2, full_text)
        for match in matches2:
            clean_match = match.strip().replace('\r', '').replace('\x07', '').replace('\n', '')
            if clean_match and any(keyword in clean_match.lower() for keyword in ['people', 'person', 'selected', 'part']):
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


def process_people_placeholders_in_document(doc):
    """
    Обробляє плейсхолдери людей у документі Word з детальною діагностикою
    """
    try:
        # Отримуємо замінники для людей
        people_replacements = people_manager.generate_replacements()

        if not people_replacements:
            #print("[DEBUG] Немає замінників для людей")
            return

        #print(f"[DEBUG] Обробляємо замінники людей: {people_replacements}")

        # Спочатку отримаємо весь текст документу для діагностики
        full_text = doc.Content.Text
        #print(f"[DEBUG] Перші 500 символів документу: {repr(full_text[:500])}")

        # Перевіряємо наявність кожного плейсхолдера в тексті
        for placeholder in people_replacements.keys():
            placeholder_exists = placeholder in full_text
            #print(f"[DEBUG] Плейсхолдер '{placeholder}' знайдено в тексті: {placeholder_exists}")
            if placeholder_exists:
                # Знаходимо контекст навколо плейсхолдера
                pos = full_text.find(placeholder)
                context_start = max(0, pos - 50)
                context_end = min(len(full_text), pos + len(placeholder) + 50)
                context = full_text[context_start:context_end]
                #print(f"[DEBUG] Контекст плейсхолдера '{placeholder}': {repr(context)}")

        # Обробляємо кожен плейсхолдер окремо
        for placeholder, replacement in people_replacements.items():
            try:
                #print(f"[DEBUG] Обробляємо плейсхолдер: '{placeholder}'")
                #print(f"[DEBUG] Заміна: '{replacement}' (довжина: {len(replacement)})")

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

        #print("[DEBUG] Обробка плейсхолдерів людей завершена")

    except Exception as e:
        print(f"[ERROR] Загальна помилка при обробці плейсхолдерів людей: {e}")
        traceback.print_exc()


def process_document_content(doc, block, current_fields):
    """
    Обробляє вміст документу Word - замінює плейсхолдери та обробляє людей
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

        # print(f"[DEBUG] Обробка всіх плейсхолдерів завершена для документу")

    except Exception as e:
        print(f"[ERROR] Помилка при обробці плейсхолдерів: {e}")


def generate_documents_word(tabview):
    selected_event = tabview.get().strip()
    current_blocks = [block for block in document_blocks if block.get("tab_name", "").strip() == selected_event]

    # print(f"[DEBUG] Вибраний захід: {selected_event}")
    # print(f"[DEBUG] Доступні блоки: {[b.get('tab_name') for b in document_blocks]}")

    if not current_blocks:
        messagebox.showwarning("Увага", f"У заході '{selected_event}' немає жодного договору для генерації.")
        return False

    save_dir = filedialog.askdirectory(title="Оберіть папку для збереження документів Word")
    if not save_dir:
        return False

    # Отримуємо динамічні поля для поточних блоків
    current_fields = get_all_placeholders_from_blocks(current_blocks)
    # print(f"[DEBUG] Динамічні поля: {current_fields}")

    # === ДОБАВЛЯЕМ ГЕНЕРАЦИЮ EXCEL В САМОМ НАЧАЛЕ ===
    excel_generated = False
    try:
        print("[INFO] Розпочинаємо генерацію Excel...")
        # Excel файл сохраняется в папке программы, а не в выбранной пользователем
        program_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(program_dir, "база_даних.xlsx")
        excel_success, excel_filename = export_document_data_to_excel(current_blocks, current_fields, excel_path, tabview)
        if excel_success:
            print(f"[INFO] Excel файл успішно згенеровано: {excel_filename}")
            excel_generated = True
        else:
            print("[WARN] Помилка при генерації Excel файлу")
    except Exception as e:
        print(f"[ERROR] Помилка при генерації Excel: {e}")

    # Зберігаємо дані для кожного блоку
    for block in current_blocks:
        if "path" in block and block["path"]:
            block_data = {}
            for field in current_fields:
                if field in block["entries"] and block["entries"][field]:
                    block_data[field] = block["entries"][field].get()
            save_memory(block_data, block["path"])

    word_app = None
    try:
        pythoncom.CoInitialize()
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False
        generated_count = 0

        for block in current_blocks:
            template_path_abs = os.path.abspath(block["path"])
            if not os.path.exists(template_path_abs):
                log_and_show_error(FileNotFoundError, f"Шаблон не знайдено: {template_path_abs}", None)
                continue

            try:
                doc = word_app.Documents.Open(template_path_abs)

                # Обробляємо звичайні плейсхолдери та людей
                process_document_content(doc, block, current_fields)

                # Обробка таблиць з товарами (якщо є items)
                items = block["items"]() if callable(block.get("items")) else block.get("items", [])

                if items:
                    for table in doc.Tables:
                        template_row = None
                        for row in table.Rows:
                            for cell in row.Cells:
                                text = cell.Range.Text.strip().replace('\r', '').replace('\x07', '')
                                if "<дк>" in text:
                                    template_row = row
                                    break
                            if template_row:
                                break

                        if not template_row:
                            continue

                        total_sum = 0
                        for i, item in enumerate(items):
                            try:
                                new_row = table.Rows.Add(template_row)
                                new_row.Cells(1).Range.Text = str(i + 1)
                                new_row.Cells(2).Range.Text = item.get("дк", "")
                                new_row.Cells(3).Range.Text = "шт."
                                new_row.Cells(4).Range.Text = item.get("кількість", "")
                                new_row.Cells(5).Range.Text = item.get("ціна за одиницю", "")

                                qty_str = item.get("кількість", "0").replace(",", ".")
                                price_str = item.get("ціна за одиницю", "0").replace(",", ".")
                                qty = float(qty_str) if qty_str else 0
                                price = float(price_str) if price_str else 0
                                suma = qty * price

                                new_row.Cells(6).Range.Text = f"{suma:.2f}"
                                total_sum += suma
                            except Exception as e:
                                print(f"[WARN] Не вдалося обробити товар: {item} → {e}")

                        try:
                            template_row.Delete()
                        except Exception as e:
                            print(f"[WARN] Не вдалося видалити шаблонний рядок: {e}")

                        # Замінюємо загальну суму
                        try:
                            find_obj = doc.Content.Find
                            find_obj.ClearFormatting()
                            find_obj.Replacement.ClearFormatting()
                            find_obj.Execute(FindText="<разом>",
                                             ReplaceWith=f"{total_sum:.2f}",
                                             Replace=win32.constants.wdReplaceAll)
                        except Exception as e:
                            print(f"[WARN] Не вдалося замінити <разом>: {e}")
                        break

                # Формуємо ім'я файлу
                base_name = os.path.splitext(os.path.basename(block["path"]).replace("ШАБЛОН", "").strip())[0]

                # Шукаємо поле для назви (товар, назва, найменування тощо)
                safe_name = "договір"
                name_fields = ["товар", "назва", "найменування", "предмет", "послуга"]
                for name_field in name_fields:
                    if name_field in block['entries'] and block['entries'][name_field]:
                        name_value = block['entries'][name_field].get().strip()
                        if name_value:
                            safe_name = "".join(c if c.isalnum() or c in " -" else "_" for c in name_value)[:50]
                            break

                save_path = os.path.join(save_dir, f"{base_name} {safe_name}.docm")

                doc.SaveAs(save_path, FileFormat=13)
                doc.Close(False)
                generated_count += 1

            except Exception as e_doc:
                log_and_show_error(type(e_doc), f"Помилка при обробці шаблону: {block['path']}\n{e_doc}",
                                   sys.exc_info()[2])
                if 'doc' in locals() and doc is not None:
                    try:
                        doc.Close(False)
                    except:
                        pass

        # === ГЕНЕРАЦИЯ КОШТОРИСА ===
        koshtorys_generated = False
        try:
            print("[INFO] Розпочинаємо генерацію кошторису...")
            from koshtorys import fill_koshtorys

            koshtorys_success = fill_koshtorys(current_blocks)
            if koshtorys_success:
                print("[INFO] Кошторис також успішно згенеровано")
                koshtorys_generated = True
            else:
                print("[WARN] Помилка при генерації кошторису")
        except ImportError:
            print("[ERROR] Не вдалося імпортувати модуль koshtorys")
        except Exception as e:
            print(f"[ERROR] Помилка при генерації кошторису: {e}")

        # === ФИНАЛЬНОЕ СООБЩЕНИЕ ===
        success_messages = []
        error_messages = []

        if generated_count > 0:
            success_messages.append(f"{generated_count} документ(и) Word")
        else:
            error_messages.append("Жодного документа Word не згенеровано")

        if excel_generated:
            success_messages.append("Excel база даних")
        else:
            error_messages.append("Excel база даних не згенеровано")

        if koshtorys_generated:
            success_messages.append("Кошторис")
        else:
            error_messages.append("Кошторис не згенеровано")

        if success_messages:
            success_text = " та ".join(success_messages)
            if error_messages:
                error_text = "\n\nПомилки: " + "; ".join(error_messages)
                excel_info = f"\n\nExcel база даних збережена в папці програми: {os.path.dirname(os.path.abspath(__file__))}" if excel_generated else ""
                messagebox.showwarning("Частковий успіх",
                                       f"Успішно згенеровано: {success_text}\n\nWord документи та кошторис збережено в: {save_dir}{excel_info}{error_text}")
            else:
                excel_info = f"\n\nExcel база даних збережена в папці програми: {os.path.dirname(os.path.abspath(__file__))}" if excel_generated else ""
                messagebox.showinfo("Повний успіх",
                                    f"Успішно згенеровано: {success_text}\n\nWord документи та кошторис збережено в: {save_dir}{excel_info}")
        else:
            messagebox.showerror("Помилка", "Нічого не вдалося згенерувати!")

        return len(success_messages) > 0

    except Exception as e_main_word:
        log_and_show_error(type(e_main_word), f"Загальна помилка при генерації документів Word: {e_main_word}",
                           sys.exc_info()[2])
        return False

    finally:
        if word_app:
            try:
                word_app.Quit()
            except:
                pass
        pythoncom.CoUninitialize()


