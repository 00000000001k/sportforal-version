# excel_export.py
from openpyxl import Workbook
import tkinter.messagebox as messagebox
import sys
# Предполагается, что error_handler.py находится в той же директории
try:
    from error_handler import log_and_show_error
except ImportError:
    # Заглушка, если error_handler не найден
    def log_and_show_error(exc_type, exc_value, exc_traceback, error_log_file="error.txt"):
        print(f"Error (logging stub): {exc_value}")
        traceback.print_exception(exc_type, exc_value, exc_traceback)
        try:
            messagebox.showerror("Ошибка (Excel Export)", f"Произошла ошибка: {str(exc_value)}")
        except: pass


# FIELDS нужно будет передавать или импортировать из основного модуля, если он константный
# Либо передавать как аргумент в функцию export_to_excel

def export_document_data_to_excel(document_blocks, fields_list, output_filename="заповнені_дані.xlsx"):
    """
    Экспортирует данные из document_blocks в Excel файл.

    Args:
        document_blocks: Список словарей, где каждый словарь представляет блок документа.
                         Каждый блок должен содержать ключ "path" (путь к шаблону)
                         и "entries" (словарь с виджетами Entry или их значениями).
        fields_list: Список строковых ключей полей, которые нужно экспортировать.
        output_filename: Имя выходного Excel файла.
    """
    try:
        wb = Workbook()
        ws = wb.active
        ws.append(["Шаблон"] + fields_list) # Заголовки столбцов

        for block in document_blocks:
            row_data = [block.get("path", "N/A")] # Путь к шаблону
            entries = block.get("entries", {})
            for field_key in fields_list:
                entry_widget_or_value = entries.get(field_key)
                if hasattr(entry_widget_or_value, 'get'): # Если это виджет с методом get()
                    row_data.append(entry_widget_or_value.get())
                else: # Если это уже значение (например, при передаче обработанных данных)
                    row_data.append(entry_widget_or_value if entry_widget_or_value is not None else "")
            ws.append(row_data)

        wb.save(output_filename)
        # messagebox.showinfo("Excel", f"Дані збережено у '{output_filename}'")
        return True
    except Exception as e:
        log_and_show_error(type(e), e, sys.exc_info()[2])
        # messagebox.showerror("Помилка Excel", f"Помилка при збереженні в Excel: {e}") # log_and_show_error уже покажет
        return False