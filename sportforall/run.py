# sportforall/run.py

import sys
import os

# Добавляем корневую папку проекта в sys.path, чтобы Python мог найти пакет sportforall
# Это нужно, если run.py запускается не из корневой папки проекта,
# но при использовании `python -m sportforall.run` из родительской папки,
# это может быть не строго необходимо, но не помешает.
# project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
# if project_root not in sys.path:
#     sys.path.insert(0, project_root)

# Устанавливаем локаль для корректной работы с символами (не всегда необходимо, но может помочь)
# import locale
# try:
#     locale.setlocale(locale.LC_ALL, 'uk_UA.UTF-8') # Пример для украинской локали
# except locale.Error:
#     try:
#         locale.setlocale(locale.LC_ALL, 'ukr_ukr') # Альтернативный вариант для Windows
#     except locale.Error:
#         print("WARNING: Could not set Ukrainian locale.")
#         pass # Не критично, если не удалось установить локаль


# Импортируем главный класс приложения из GUI модуля
# !!! ИСПРАВЛЕНО: Импортируем MainApplication напрямую из gui.main_app !!!
# Было: from sportforall.data_manager import load_data, save_data # Эту строку удалили
# Было: from sportforall.gui.main_app import MainApp # Возможно старое название класса
from sportforall.gui.main_app import MainApplication # Правильный импорт MainApplication из main_app.py

# Импортируем модуль логирования ошибок, чтобы он был доступен
from sportforall import error_handling

def main():
    """Основная точка входа в приложение."""
    print("DEBUG RUN: Starting application main function.") # Для дебагу
    # Код инициализации или запуска основного окна GUI
    try:
        app = MainApplication()
        app.mainloop()
    except Exception as e:
        print(f"ERROR RUN: An unhandled error occurred in the main application loop: {e}")
        # Логируем неперехваченную ошибку
        error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Неперехоплена помилка в основному циклі додатку: {e}")
        # TODO: Возможно, показать пользователю сообщение об ошибке


if __name__ == "__main__":
    # Устанавливаем режим запуска модуля для корректного обнаружения пакетов
    # sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    # print(f"DEBUG RUN: sys.path: {sys.path}") # Для отладки путей
    main()