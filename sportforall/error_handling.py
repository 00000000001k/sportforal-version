# sportforall/error_handling.py

import os
import datetime
import traceback
import sys

def log_error(exc_type, exc_value, exc_traceback, level="ERROR", message=None):
    """
    Логирует информацию об ошибке в файл.
    Добавлено принудительное создание директории и отладочные принты для диагностики проблем с логированием.

    Args:
        exc_type: Тип исключения (например, ValueError). Может быть None.
        exc_value: Значение исключения (экземпляр исключения). Может быть None.
        exc_traceback: Объект трассировки (из sys.exc_info()[2]). Может быть None.
        level: Уровень логирования (например, "INFO", "WARNING", "ERROR", "CRITICAL").
        message: Дополнительное сообщение об ошибке.
    """
    log_dir = "logs"
    log_file_path = os.path.join(log_dir, "application.log")

    # !!! ДОБАВЛЕНО: Принудительное создание директории для логов с отладкой !!!
    try:
        if not os.path.exists(log_dir):
            print(f"DEBUG LOGGING: Directory '{log_dir}' does not exist. Attempting to create...") # Отладочный вывод
            os.makedirs(log_dir)
            print(f"DEBUG LOGGING: Directory '{log_dir}' created successfully.") # Отладочный вывод
        else:
            print(f"DEBUG LOGGING: Directory '{log_dir}' already exists.") # Отладочный вывод
    except OSError as e:
        # Если не удалось создать директорию, выводим критическую ошибку в консоль и детали исходной ошибки
        print(f"FATAL ERROR LOGGING: Failed to create directory '{os.path.abspath(log_dir)}': {e}") # Критическая ошибка в консоль
        # Пытаемся вывести детали исходной ошибки, которую хотели залогировать
        print("Original error details (directory creation failed):")
        timestamp_console = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        console_message = f"[{timestamp_console}] [{level.upper()}] "
        if message:
             console_message += f"Message: {message}\n"
        if exc_type is not None:
             console_message += f"Exception Type: {exc_type.__name__}, Value: {exc_value}\n"
             if exc_traceback is not None:
                  try:
                       # Попытка форматировать трассировку для консоли
                       console_message += "Traceback:\n" + "".join(traceback.format_exception(exc_type, exc_value, exc_traceback)) + "\n"
                  except Exception as format_exc:
                       console_message += f"Error formatting traceback for console: {format_exc}\n"
        print(console_message)
        # Не можем записать в файл, поэтому просто возвращаемся
        return


    # Формируем сообщение для записи в лог-файл
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log_message = f"[{timestamp}] [{level.upper()}] "

    # Добавляем дополнительное сообщение, если оно есть
    if message:
        log_message += f"Message: {message}\n"

    # Добавляем информацию об исключении и трассировку, если они есть
    if exc_type is not None:
         # Если сообщения не было, добавляем информацию об исключении в первую строку
         if not message:
             log_message += f"Exception Type: {exc_type.__name__}, Value: {exc_value}\n"
         else:
             # Если сообщение было, добавляем информацию об исключении в новой строке
             log_message += f"Exception Type: {exc_type.__name__}, Value: {exc_value}\n"


         # Добавляем информацию о трассировке стека
         if exc_traceback is not None:
            try:
                # Используем traceback.format_exception для форматирования трассировки
                traceback_str_list = traceback.format_exception(exc_type, exc_value, exc_traceback)
                traceback_str = "".join(traceback_str_list)
                log_message += "Traceback:\n" + traceback_str + "\n"

            except Exception as format_exc:
                 # Если форматирование трассировки не удалось
                 log_message += f"Ошибка форматирования трассировки: {format_exc}\n"
                 # Попробуем хотя бы строку из traceback
                 try:
                      log_message += f"Traceback object: {exc_traceback}\n"
                 except Exception as tb_str_exc:
                      log_message += "Не удалось получить строку из объекта трассировки.\n"


    log_message += "="*50 + "\n" # Добавляем разделитель


    # Записываем сообщение в лог-файл
    try:
        print(f"DEBUG LOGGING: Attempting to write log to '{os.path.abspath(log_file_path)}'...") # Отладочный вывод
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(log_message)
        print(f"DEBUG LOGGING: Log successfully written to '{os.path.abspath(log_file_path)}'") # Отладочный вывод

    except Exception as e:
        # Если не удалось записать в лог-файл, печатаем ошибку в консоль
        print(f"FATAL ERROR: Не удалось записать лог ошибки в файл {os.path.abspath(log_file_path)}. Ошибка: {e}")
        print("Исходная ошибка (file writing failed):")
        print(log_message) # Печатаем исходное сообщение, которое хотели записать
        traceback.print_exc() # Печатаем трассировку ошибки записи в лог-файл