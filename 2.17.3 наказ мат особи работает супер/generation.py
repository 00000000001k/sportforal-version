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


def extract_placeholders_from_word(template_path):
    """
    Витягує всі плейсхолдери типу <поле> з документу Word
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

        # Також перевіряємо таблиці окремо
        for table in doc.Tables:
            for row in table.Rows:
                for cell in row.Cells:
                    full_text += " " + cell.Range.Text

        # Шукаємо всі плейсхолдери типу <поле>
        pattern = r'<([^>]+)>'
        matches = re.findall(pattern, full_text)

        for match in matches:
            # Очищуємо від спеціальних символів Word
            clean_match = match.strip().replace('\r', '').replace('\x07', '')
            if clean_match:
                placeholders.add(clean_match)

        print(f"[DEBUG] Знайдені плейсхолдери в {template_path}: {placeholders}")

    except Exception as e:
        print(f"[ERROR] Помилка при витягуванні плейсхолдерів з {template_path}: {e}")
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
            block_placeholders = extract_placeholders_from_word(block["path"])
            all_placeholders.update(block_placeholders)

            # Зберігаємо плейсхолдери для кожного блоку
            block["placeholders"] = block_placeholders

    return sorted(list(all_placeholders))


def process_people_placeholders_in_document(doc):
    """
    Обробляє плейсхолдери людей у документі Word
    """
    try:
        # Отримуємо замінники для людей
        people_replacements = people_manager.generate_replacements()

        if not people_replacements:
            print("[DEBUG] Немає замінників для людей")
            return

        print(f"[DEBUG] Обробляємо замінники людей: {people_replacements}")

        # Обробляємо кожен плейсхолдер окремо
        for placeholder, replacement in people_replacements.items():
            try:
                # Якщо replacement порожній - видаляємо весь абзац
                if replacement == "":
                    # Шукаємо абзац з плейсхолдером і видаляємо його повністю
                    find_obj = doc.Content.Find
                    find_obj.ClearFormatting()

                    while find_obj.Execute(FindText=placeholder):
                        # Розширюємо виділення до всього абзацу
                        paragraph = find_obj.Parent.Paragraphs(1)
                        paragraph.Range.Delete()
                        # Перериваємо цикл, оскільки абзац видалено
                        break
                else:
                    # Звичайна заміна для обраних людей
                    find_obj = doc.Content.Find
                    find_obj.ClearFormatting()
                    find_obj.Replacement.ClearFormatting()

                    # Специальная обработка для многострочного текста
                    if "\r\n" in replacement:
                        word_replacement = replacement.replace("\r\n", "^p")
                    else:
                        word_replacement = replacement

                    # Замінюємо плейсхолдер на заміну
                    result = find_obj.Execute(
                        FindText=placeholder,
                        ReplaceWith=word_replacement,
                        Replace=win32.constants.wdReplaceAll
                    )

                    if result:
                        print(f"[DEBUG] Замінено {placeholder} -> {replacement[:50]}...")

            except Exception as e:
                print(f"[ERROR] Помилка при заміні {placeholder}: {e}")

        # Також обробляємо таблиці окремо
        for table in doc.Tables:
            for row in table.Rows:
                for cell in row.Cells:
                    try:
                        cell_text = cell.Range.Text
                        modified = False

                        # Перевіряємо кожен плейсхолдер у тексті комірки
                        for placeholder, replacement in people_replacements.items():
                            if placeholder in cell_text:
                                # Для таблиц также обрабатываем переносы строк
                                if replacement == "":
                                    # Видаляємо плейсхолдер з комірки
                                    cell_text = cell_text.replace(placeholder, "")
                                else:
                                    if "\r\n" in replacement:
                                        word_replacement = replacement.replace("\r\n", "\r")
                                    else:
                                        word_replacement = replacement
                                    cell_text = cell_text.replace(placeholder, word_replacement)
                                modified = True

                        # Якщо текст змінився, оновлюємо комірку
                        if modified:
                            cell.Range.Text = cell_text

                    except Exception as e:
                        print(f"[ERROR] Помилка при обробці комірки таблиці: {e}")

        print("[DEBUG] Обробка плейсхолдерів людей завершена")

    except Exception as e:
        print(f"[ERROR] Загальна помилка при обробці плейсхолдерів людей: {e}")


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

        print(f"[DEBUG] Обробка всіх плейсхолдерів завершена для документу")

    except Exception as e:
        print(f"[ERROR] Помилка при обробці плейсхолдерів: {e}")


def generate_documents_word(tabview):
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

        if generated_count > 0:
            messagebox.showinfo("Успіх", f"{generated_count} документ(и) Word збережено успішно в папці:\n{save_dir}")
        else:
            messagebox.showwarning("Увага", "Жодного документа Word не було згенеровано через помилки.")
        return True

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


def combined_generation_process(tabview):
    if not document_blocks:
        messagebox.showwarning("Увага", "Не додано жодного договору для генерації.")
        return

    # Отримуємо всі динамічні поля
    current_event = tabview.get()
    relevant_blocks = [b for b in document_blocks if b.get("tab_name") == current_event]
    dynamic_fields = get_all_placeholders_from_blocks(relevant_blocks)

    print(f"[DEBUG] Перевіряємо поля: {dynamic_fields}")

    # Перевіряємо заповненість полів
    for i, block in enumerate(relevant_blocks, start=1):
        block_placeholders = block.get("placeholders", dynamic_fields)
        for field in block_placeholders:
            entry_widget = block["entries"].get(field)

            # Пропускаємо readonly поля
            if entry_widget and hasattr(entry_widget, 'cget') and entry_widget.cget("state") == "readonly":
                continue

            if not entry_widget or not entry_widget.get().strip():
                messagebox.showerror("Помилка заповнення", f"Блок договору #{i}: поле <{field}> порожнє.")
                return

    # Експортуємо в Excel з динамічними полями
    if not export_document_data_to_excel(document_blocks, dynamic_fields):
        messagebox.showerror("Помилка", "Не вдалося згенерувати Excel файл з даними договорів.")
        return

    if not generate_documents_word(tabview):
        messagebox.showwarning("Увага", "Генерація документів Word не була повністю успішною або скасована.")

    if koshtorys.fill_koshtorys(document_blocks):
        messagebox.showinfo("Завершено", "Усі документи (Excel, Word, Кошторис) оброблено.")
    else:
        messagebox.showerror("Помилка Кошторису", "Не вдалося заповнити файл кошторису.")