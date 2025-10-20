# sportforall/word_generator_win32.py

import win32com.client
import os
import sys
import traceback
import ctypes # Додано для використання GetShortPathName

# Додано для використання в log_error
from sportforall import error_handling
from sportforall.models import Contract # Импорт модели Contract для типизации
# import messagebox # Возможно потребуется, если будем показывать ошибки здесь


# Численные значения констант Word для надежности
# wdOpenFormatXMLMacroEnabled = 13 # Формат .docm
# wdFormatDocumentDefault = 16 # Формат .docx
# wdFormatRTF = 6 # Формат .rtf
# wdFormatPDF = 17 # Формат .pdf

# Word constants for Find.Execute - using integer values for reliability
# These might need adjustment depending on Word version, but usually stable
wdFindContinue = 1
wdReplaceAll = 2


# --- Функция для получения короткого имени файла (8.3 формат) ---
# Используем ctypes для вызова GetShortPathName Windows API
def get_short_path_name(long_path):
    """Gets the short path name for a path using Windows API."""
    # print(f"DEBUG SHORT_PATH: Attempting to get short path for: {long_path}") # For debug
    if not long_path:
        # print("DEBUG SHORT_PATH: Input path is empty.") # For debug
        return long_path # Return empty if input is empty

    # Windows API GetShortPathNameW (для Unicode)
    _GetShortPathName = ctypes.windll.kernel32.GetShortPathNameW # Correct function name
    buffer_size = 260 # Максимальная длина пути
    buffer = ctypes.create_unicode_buffer(buffer_size)

    # Проверяем существование файла/директории перед вызовом API
    # GetShortPathName требует, чтобы путь существовал
    if not os.path.exists(long_path):
        print(f"WARNING SHORT_PATH: Path does not exist, cannot get short path: {long_path}") # For debug
        # Логируем предупреждение без исключения
        error_handling.log_error(None, None, None, level="WARNING", message=f"Не існує шлях для отримання короткого імені: {long_path}")
        return long_path # Возвращаем оригинальный путь, возможно с ним сработает


    # !!! ИСПРАВЛЕНО: Исправлена опечатка в имени функции - убрано лишнее "Path" !!!
    # Было: result = _GetShortPathPathName(long_path, buffer, buffer_size)
    result = _GetShortPathName(long_path, buffer, buffer_size) # Правильный вызов


    if result == 0:
        # Произошла ошибка при вызове API
        last_error = ctypes.get_last_error()
        print(f"ERROR SHORT_PATH: Failed to get short path for '{long_path}'. WinAPI error code: {last_error}") # For debug
        # Логируем ошибку вызова API
        error_handling.log_error(None, None, None, level="ERROR", message=f"Помилка WinAPI GetShortPathName для шляху '{long_path}'. Код: {last_error}")
        return long_path # Возвращаем оригинальный путь на случай ошибки API

    # print(f"DEBUG SHORT_PATH: Short path obtained: {buffer.value}") # For debug
    return buffer.value # Возвращаем полученный короткий путь


def generate_document_win32(contract: Contract, template_path: str, output_dir: str) -> str | None:
    """
    Генерує документ Word на основі шаблону, заповнюючи його даними договору.
    Використовує win32com.client (тільки для Windows).
    Пытается использовать короткие имена файлов для шаблонов с не-ASCII символами.
    Формирует имя файла из имени шаблона (без "ШАБЛОН") + поле "товар".
    Заменяет закладки и ТЕКСТОВЫЕ плейсхолдеры (<название_поля>).

    Args:
        contract: Об'єкт договору (Contract).
        template_path: Повний шлях до файлу шаблону .docm.
        output_dir: Папка для збереження згенерованого файлу.

    Returns:
        Повний шлях до згенерованого файлу, або None у разі помилки.
    """
    # !!! Отладочные сообщения о полученных путях !!!
    print(f"DEBUG GENERATOR: Received template_path: '{template_path}')") # Используем кавычки для наглядности
    print(f"DEBUG GENERATOR: Received output_dir: '{output_dir}')") # Используем кавычки для наглядности
    print(f"DEBUG GENERATOR: Received contract object: {contract})") # Debug print the contract object itself


    word = None
    doc = None
    generated_file_path = None
    try:
        print("DEBUG GENERATOR: Inside try block...") # Для дебагу

        # --- Шаг 1: Проверка шаблона и получение короткого пути ---
        print(f"DEBUG GENERATOR: Checking if template file exists at '{template_path}'...") # Для дебагу
        if not os.path.exists(template_path):
            error_message = f"Файл шаблону не знайдено: {template_path}"
            print(f"DEBUG GENERATOR: Template file check failed: {error_message}") # Отладочное сообщение о проверке
            error_handling.log_error(None, None, None, level="ERROR", message=error_message)
            # TODO: Показать messagebox пользователю через колбек в main_app
            return None
        else:
             print(f"DEBUG GENERATOR: Template file exists at '{template_path}')") # Отладочное сообщение о проверке

        # Получаем короткий путь к шаблону для надежности при работе с win32com
        template_path_for_word = get_short_path_name(template_path)
        print(f"DEBUG GENERATOR: Using template_path for Word (short/original): '{template_path_for_word}')") # Отладочное сообщение о пути для Word


        # --- Шаг 2: Проверка и создание папки сохранения ---
        print("DEBUG GENERATOR: Checking/creating output directory...") # Для дебагу
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except OSError as e:
                error_message = f"Не вдалося створити теку для збереження '{output_dir}': {e}"
                print(f"DEBUG GENERATOR: Failed to create output directory: {error_message}") # Отладочный вывод
                error_handling.log_error(type(e), e, sys.exc_info()[2], level="ERROR", message=error_message)
                # TODO: Показать messagebox пользователю через колбек в main_app
                return None


        # --- Шаг 3: Запускаем Word и открываем шаблон ---
        print("DEBUG GENERATOR: Starting Word...") # Для дебагу
        word = win32com.client.Dispatch("Word.Application")
        # word.Visible = True # Для отладки: сделать Word видимым

        print(f"DEBUG GENERATOR: Attempting to open template in Word: '{template_path_for_word}')") # Используем КОРОТКИЙ или ОРИГИНАЛЬНЫЙ путь
        doc = word.Documents.Open(template_path_for_word) # Открываем шаблон по короткому пути (или оригинальному если не сработало)
        print("DEBUG GENERATOR: Template opened successfully.") # Для дебагу


        # --- Шаг 4: Заполнение полей договора (закладки и ТЕКСТОВЫЕ плейсхолдеры) ---
        print("DEBUG GENERATOR: Filling document fields (Bookmarks and Text Placeholders)...") # Для дебагу

        # Итерация по всем полям договора для замены
        for field_name, field_value in contract.fields.items():
             # Значение для замены должно быть строкой
             replace_value = str(field_value)

             # 4.1: Заполнение закладок (Bookmark) - существующая логика
             # print(f"DEBUG GENERATOR: Attempting to replace Bookmark: {field_name} with '{replace_value}'") # For debug
             bookmark_name = field_name # Assuming bookmark name is the same as field_name
             if doc.Bookmarks.Exists(bookmark_name):
                  try:
                       rng = doc.Bookmarks(bookmark_name).Range
                       rng.Text = replace_value # Устанавливаем текст закладки
                       # print(f"DEBUG GENERATOR: Bookmark '{bookmark_name}' replaced.") # For debug
                  except Exception as fill_bookmark_error:
                       print(f"DEBUG GENERATOR: Помилка заповнення закладки '{field_name}': {fill_bookmark_error}") # Для дебагу
                       error_handling.log_error(type(fill_bookmark_error), fill_bookmark_error, sys.exc_info()[2], level="WARNING", message=f"Помилка заповнення закладки '{field_name}' в шаблоні.")
             # else:
                  # print(f"DEBUG GENERATOR: Bookmark '{field_name}' not found.") # For debug


             # 4.2: Замена текстовых плейсхолдеров (<название_поля>)
             # Формируем строку для поиска (например, "<назва_поля>")
             search_text_placeholder = f"<{field_name}>"
             # print(f"DEBUG GENERATOR: Attempting to replace Text Placeholder: {search_text_placeholder} with '{replace_value}'") # For debug

             # Используем Find.Execute для поиска и замены по всему документу
             # Необходимо выполнить поиск во всех "Историях" документа (основной текст, колонтитулы и т.д.)
             for story_range in doc.StoryRanges:
                  # print(f"DEBUG GENERATOR: Searching in Story Range type: {story_range.StoryType}") # For debug
                  # Reset Find options before each search
                  find_obj = story_range.Find
                  find_obj.ClearFormatting()
                  find_obj.Replacement.ClearFormatting()

                  # Set find parameters
                  find_obj.Text = search_text_placeholder
                  find_obj.Replacement.Text = replace_value
                  find_obj.Forward = True # Искать вперед
                  find_obj.Wrap = wdFindContinue # Продолжать поиск после достижения конца части истории
                  find_obj.Format = False # Не искать форматирование
                  
                  # Determine MatchWholeWord based on whether it's a <...> placeholder
                  # We want to find exactly <name>, not <name> in <name_suffix>
                  # But we might want to find "товар" if field_name is "товар" and placeholder is "товар"
                  # Let's make it MatchWholeWord=True if the search string starts with < and ends with >, otherwise False
                  find_whole_word = False
                  if search_text_placeholder.startswith('<') and search_text_placeholder.endswith('>'):
                      find_whole_word = True

                  find_obj.MatchWholeWord = find_whole_word # Искать как целое слово или как часть
                  find_obj.MatchCase = False # Игнорировать регистр


                  find_obj.MatchWildcards = False # Не использовать подстановочные знаки
                  find_obj.MatchSoundsLike = False # Не использовать "похожее звучание"
                  find_obj.MatchAllWordForms = False # Не использовать все формы слова

                  try:
                      # Execute the find and replace operation
                      # wdReplaceOne = 1 to replace one at a time and loop
                      # This is generally safer than wdReplaceAll on the whole document range
                      while find_obj.Execute(Replace=win32com.client.constants.wdReplaceOne, Wrap=wdFindContinue):
                          pass # Loop continues as long as matches are found and replaced

                      # print(f"DEBUG GENERATOR: Text Placeholder '{search_text_placeholder}' replacement attempted.") # For debug

                  except Exception as replace_error:
                      print(f"DEBUG GENERATOR: Error replacing text '{search_text_placeholder}': {replace_error}") # For debug
                      # Log the error with more context
                      error_handling.log_error(type(replace_error), replace_error, sys.exc_info()[2], level="WARNING", message=f"Помилка заміни тексту '{search_text_placeholder}' в шаблоні.")
                      pass # Continue with other fields

        # TODO: Also handle MergeFields if used in the template format


        # --- Шаг 5: Формируем имя файла для сохранения (пользовательская логика) ---
        print("DEBUG GENERATOR: Forming save file name...") # Для дебагу
        # ... (Your existing code for forming save_path based on template name and 'товар' field) ...
        # Получаем полное имя файла шаблона без пути
        template_full_name = os.path.basename(template_path)
        print(f"DEBUG GENERATOR: Template full name: '{template_full_name}')") # Для дебагу

        # Получаем базовое имя файла шаблона без расширения и без пути
        template_base_name = os.path.splitext(template_full_name)[0]
        print(f"DEBUG GENERATOR: Template base name: '{template_base_name}')") # Для дебагу


        # Убираем "ШАБЛОН " из начала имени, если оно там есть (независимо от регистра)
        if template_base_name.lower().startswith("шаблон "):
             template_base_name = template_base_name[len("ШАБЛОН "):]
             print(f"DEBUG GENERATOR: Template base name after removing 'ШАБЛОН ': '{template_base_name}')") # Для дебагу

        # Получаем значение поля "товар" из договора
        # TODO: Убедиться, что поле "товар" существует в constants.FIELDS
        # TODO: Возможно, нужно использовать id поля вместо имени "товар" для надежности
        # Получаем значение из словаря fields объекта contract
        item_field_value = contract.fields.get("товар", "") # Получаем значение поля "товар" (по имени)
        print(f"DEBUG GENERATOR: Raw value of field 'товар': '{item_field_value}')") # Для дебагу


        # Санируем значение поля "товар" для использования в имени файла
        safe_item_field_value = "".join([c for c in str(item_field_value) if c.isalnum() or c in (' ', '-', '_')]).rstrip()
        print(f"DEBUG GENERATOR: Sanitized value of field 'товар': '{safe_item_field_value}')") # Для дебагу


        # Составляем конечное имя файла
        final_name_parts = [template_base_name] # Начинаем с имени шаблона без "ШАБЛОН"
        if safe_item_field_value:
             final_name_parts.append(safe_item_field_value) # Добавляем санированное значение поля "товар"

        # Объединяем части имени через пробел и добавляем расширение .docm
        safe_contract_name_base = " ".join(final_name_parts).strip() # Объединяем части, убираем лишние пробелы

        # Санируем конечное имя файла еще раз на всякий случай
        # Разрешаем буквы, цифры, пробелы, точки, тире, подчеркивания
        safe_contract_name_final = "".join([c for c in safe_contract_name_base if c.isalnum() or c in (' ', '.', '-', '_')]).rstrip()

        # Убедимся, что имя не пустое
        if not safe_contract_name_final:
             safe_contract_name_final = f"Договір_{contract.id[:8]}" # Используем часть ID, если имя пустое

        file_name = f"{safe_contract_name_final}.docm" # Добавляем расширение .docm
        print(f"DEBUG GENERATOR: Final file_name formed: '{file_name}')") # Для дебагу


        # Формируем полный путь для сохранения
        save_path = os.path.join(output_dir, file_name) # <-- save_path для сохранения
        print(f"DEBUG GENERATOR: Final save_path formed: '{save_path}')") # Для дебагу

        # --- Шаг 6: Сохранение документа ---
        print(f"DEBUG GENERATOR: Saving document to '{save_path}'...") # Для дебагу
        # При сохранении тоже можно попробовать использовать короткий путь к папке, если проблемы с записью
        # save_path_for_word = get_short_path_name(save_path)
        # doc.SaveAs(save_path_for_word, FileFormat=13) # Или использовать короткий путь
        doc.SaveAs(save_path, FileFormat=13) # Используем полный путь для сохранения (Windows обычно с этим справляется)
        print("DEBUG GENERATOR: Document saved successfully.") # Для дебагу

        generated_file_path = save_path
        print(f"DEBUG GENERATOR: Document successfully generated: {generated_file_path}") # Для дебагу


        # --- Шаг 7: Закрытие документа и Word (ВНУТРИ TRY) ---
        print("DEBUG GENERATOR: Closing document...") # Для дебагу
        doc.Close(SaveChanges=0) # 0 = wdDoNotSaveChanges
        print("DEBUG GENERATOR: Document closed.") # Для дебагу

        print("DEBUG GENERATOR: Quitting Word...") # Для дебагу
        word.Quit()
        print("DEBUG GENERATOR: Word closed.") # Для дебагу

        return generated_file_path # Успех, возвращаем путь к сгенерированному файлу


    except Exception as e:
        # !!! СТАНДАРТНЫЙ EXCEPT БЛОК С ЛОГИРОВАНИЕМ ВОССТАНОВЛЕН !!!
        # Логируем и обрабатываем ошибку
        error_message = f"Помилка генерації документа для договору '{contract.name}' з шаблону '{os.path.basename(template_path)}': {e}"
        print(f"ERROR during document generation: {error_message}") # Печатаем краткое сообщение

        # Cleanup Word/doc (попробуем закрыть даже в случае ошибки, с логированием ошибок закрытия)
        if doc:
            try:
                doc.Close(SaveChanges=0) # 0 = wdDoNotSaveChanges
            except Exception as close_error:
                print(f"DEBUG GENERATOR: Error closing document after generation error: {close_error}")
                error_handling.log_error(type(close_error), close_error, sys.exc_info()[2], level="WARNING", message="Помилка при закритті документа після помилки генерації.")
            doc = None

        if word:
            try:
                word.Quit()
            except Exception as quit_error:
                print(f"DEBUG GENERATOR: Error quitting Word after generation error: {quit_error}")
                # !!! ИСПРАВЛЕНО: Добавлены недостающие аргументы в вызов log_error !!!
                error_handling.log_error(type(quit_error), quit_error, sys.exc_info()[2], level="WARNING", message="Помилка при закритті Word після помишки генерації.", exc_info=True) # Добавлено exc_info=True для полной трассировки
            word = None


        # Получаем информацию об исключении для логирования
        exc_type, exc_value, exc_traceback = sys.exc_info()

        # Логируем ошибку
        error_handling.log_error(exc_type, exc_value, exc_traceback, message=error_message)


        return None # Return None on error

    # block finally удален