# generation.py
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import tkinter.messagebox as messagebox
from tkinter import filedialog
import pythoncom
import win32com.client as win32

from globals import FIELDS, document_blocks
from data_persistence import save_memory
from excel_export import export_document_data_to_excel
from text_utils import number_to_ukrainian_text
import koshtorys
from error_handler import log_and_show_error


def generate_documents_word(tabview):
    selected_event = tabview.get().strip()  # Убираем .lower()
    current_blocks = [block for block in document_blocks if block.get("tab_name", "").strip() == selected_event]

    print(f"[DEBUG] Вибраний захід: {selected_event}")
    print(f"[DEBUG] Доступні блоки: {[b.get('tab_name') for b in document_blocks]}")

    if not current_blocks:
        messagebox.showwarning("Увага", f"У заході '{selected_event}' немає жодного договору для генерації.")
        return False

    save_dir = filedialog.askdirectory(title="Оберіть папку для збереження документів Word")
    if not save_dir:
        return False

    current_event = tabview.get()
    relevant_blocks = [b for b in document_blocks if b.get("tab_name") == current_event]

    for block in current_blocks:
        if "path" in block and block["path"]:
            block_data = {f: block["entries"][f].get() for f in FIELDS if f in block["entries"]}
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

                for key in FIELDS:
                    if key in block["entries"]:
                        find_obj = doc.Content.Find
                        find_obj.ClearFormatting()
                        find_obj.Replacement.ClearFormatting()
                        find_obj.Execute(FindText=f"<{key}>",
                                         ReplaceWith=block["entries"][key].get(),
                                         Replace=win32.constants.wdReplaceAll)

                items = block["items"]() if callable(block.get("items")) else block.get("items", [])

                if items:
                    for table in doc.Tables:
                        for row in table.Rows:
                            for cell in row.Cells:
                                text = cell.Range.Text.strip().replace('\r', '').replace('\x07', '')
                                if "<дк>" in text:
                                    template_row = row
                                    break
                            else:
                                continue
                            break
                        else:
                            continue

                        total_sum = 0
                        for i, item in enumerate(items):
                            try:
                                new_row = table.Rows.Add(template_row)
                                new_row.Cells(1).Range.Text = str(i + 1)
                                new_row.Cells(2).Range.Text = item["дк"]
                                new_row.Cells(3).Range.Text = "шт."
                                new_row.Cells(4).Range.Text = item["кількість"]
                                new_row.Cells(5).Range.Text = item["ціна за одиницю"]
                                qty = float(item["кількість"].replace(",", "."))
                                price = float(item["ціна за одиницю"].replace(",", "."))
                                suma = qty * price
                                new_row.Cells(6).Range.Text = f"{suma:.2f}"
                                total_sum += suma
                            except Exception as e:
                                print(f"[WARN] Не вдалося обробити товар: {item} → {e}")

                        try:
                            template_row.Delete()
                        except Exception as e:
                            print(f"[WARN] Не вдалося видалити шаблонний рядок: {e}")

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

                base_name = os.path.splitext(os.path.basename(block["path"]).replace("ШАБЛОН", "").strip())[0]
                товар_name = block['entries'].get('товар', None)
                safe_name = "".join(c if c.isalnum() or c in " -" else "_" for c in товар_name.get())[
                            :50] if товар_name else "договір"
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

    for i, block in enumerate(document_blocks, start=1):
        for field in FIELDS:
            entry_widget = block["entries"].get(field)
            if field in ["сума прописом", "разом"] and entry_widget and entry_widget.cget("state") == "readonly":
                continue
            if not entry_widget or not entry_widget.get().strip():
                messagebox.showerror("Помилка заповнення", f"Блок договору #{i}: поле <{field}> порожнє.")
                return

    if not export_document_data_to_excel(document_blocks, FIELDS):
        messagebox.showerror("Помилка", "Не вдалося згенерувати Excel файл з даними договорів.")
        return

    if not generate_documents_word(tabview):
        messagebox.showwarning("Увага", "Генерація документів Word не була повністю успішною або скасована.")

    if koshtorys.fill_koshtorys(document_blocks):
        messagebox.showinfo("Завершено", "Усі документи (Excel, Word, Кошторис) оброблено.")
    else:
        messagebox.showerror("Помилка Кошторису", "Не вдалося заповнити файл кошторису.")
