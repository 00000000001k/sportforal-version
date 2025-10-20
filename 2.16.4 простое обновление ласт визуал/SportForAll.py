# sport.py

import tkinter as tk

import tkinter.messagebox as messagebox

import traceback

import sys



# Импортируем разделенные модули
try:
    from error_handler import log_and_show_error, setup_global_exception_handler
    from gui_utils import SafeCTk, bind_entry_shortcuts, create_context_menu
    from custom_widgets import CustomEntry
    from auth_utils import ask_password  # APP_PASSWORD теперь внутри auth_utils
    from data_persistence import load_memory, save_memory, get_template_memory, MEMORY_FILE
    from excel_export import export_document_data_to_excel
    from text_utils import number_to_ukrainian_text  # Нужен для авто-заполнения "сума прописом"
    import koshtorys  # Для вызова fill_koshtorys и доступа к его настройкам

    from globals import FIELDS, EXAMPLES, version, document_blocks
    from generation import generate_documents_word, combined_generation_process

except ImportError as e:
    # Это критическая ошибка, приложение не сможет работать
    error_message = (f"Критична помилка імпорту в sport.py: {e}.\n"
                     f"Переконайтеся, що всі файли (.py) знаходяться в одній директорії.\n"
                     f"Трасування: {traceback.format_exc()}")
    print(error_message)
    # Попытка показать messagebox, если Tkinter инициализирован
    try:
        root_temp = tk.Tk()
        root_temp.withdraw()  # Скрыть временное окно
        messagebox.showerror("Критична помилка імпорту", error_message)
        root_temp.destroy()
    except:
        pass  # Если GUI не работает, сообщение уже выведено в консоль
    sys.exit(1)

# Устанавливаем глобальный обработчик ошибок
setup_global_exception_handler()

# Попытка импорта win32com
try:
    import win32com.client as win32
except ImportError:
    log_and_show_error(ImportError,
                       "Не вдалося імпортувати модуль win32com.client.\nУстановіть його командою: pip install pywin32",
                       None)
    sys.exit(1)

# Глобальные переменные для GUI основного приложения
main_app_root = None
scroll_frame_main = None
main_context_menu = None
tabview = None


def make_show_hint_command(hint_text, field_name):
    """Создает команду для кнопки подсказки, отображающую messagebox."""

    def show_hint():
        messagebox.showinfo(f"Підказка для <{field_name}>", hint_text)

    return show_hint


def on_app_close():
    global main_app_root, tabview
    try:
        from data_persistence import save_current_state
        save_current_state(tabview, document_blocks)

    except Exception as e:
        from data_persistence import log_and_show_error
        log_and_show_error(type(e), e, sys.exc_info()[2])
    finally:
        main_app_root.destroy()


# ---------------- ОСНОВНИЙ ІНТЕРФЕЙС ----------------
from auth_utils import ask_password
from error_handler import log_and_show_error
from app import launch_main_app

# --- Точка входа в программу ---
if __name__ == "__main__":
    try:
        # Імпорт splash screen з паролем
        from splash_screen import show_splash_with_password


        def on_success():
            global tabview, main_app_root

            main_app_root, tabview = launch_main_app()

            # Запускаємо перевірку оновлень через after, щоб mainloop вже працював
            from ctk_update_manager import setup_auto_updates
            update_manager = setup_auto_updates(main_app_root, version)

            main_app_root.after(1000, update_manager.check_updates_manual)  # вручну або авто



            main_app_root.mainloop()


        # Показуємо splash screen з інтегрованим паролем
        # Пароль береться з auth_utils
        from auth_utils import APP_PASSWORD

        show_splash_with_password(
            duration=4,
            app_password=APP_PASSWORD,
            success_callback=on_success
        )

    except Exception as e_main_start:
        log_and_show_error(
            type(e_main_start),
            f"Критична помилка при старті програми: {e_main_start}",
            sys.exc_info()[2]
        )
        input("Натисніть Enter для завершення...")
        sys.exit(1)