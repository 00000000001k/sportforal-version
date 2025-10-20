import tkinter.messagebox as messagebox
import traceback
import datetime
import sys

ERROR_LOG = "error.txt"

def log_error_to_file(exc_type, exc_value, exc_traceback, error_log_file=ERROR_LOG):
    """Записывает информацию об ошибке в лог-файл."""
    try:
        with open(error_log_file, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
            traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)
    except Exception as ex:
        print(f"Ошибка при логировании в файл: {str(ex)}")

def log_and_show_error(exc_type, exc_value, exc_traceback, error_log_file=ERROR_LOG):
    """Логирует ошибку и показывает сообщение пользователю."""
    log_error_to_file(exc_type, exc_value, exc_traceback, error_log_file)
    try:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {str(exc_value)}\nПодробности в файле {error_log_file}")
    except Exception as mb_ex:
        print(f"Ошибка при показе messagebox: {mb_ex}")
        # Если GUI недоступен, просто выводим в консоль
        print(f"Произошла ошибка: {str(exc_value)}\nПодробности в файле {error_log_file}")


def setup_global_exception_handler():
    """Устанавливает глобальный обработчик исключений."""
    sys.excepthook = log_and_show_error

# Для использования в koshtorys.py
def log_error_koshtorys(exc_type, exc_value, exc_traceback, error_log="error.txt"):
    log_and_show_error(exc_type, exc_value, exc_traceback, error_log_file=error_log)