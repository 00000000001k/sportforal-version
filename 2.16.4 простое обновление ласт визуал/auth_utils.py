# auth_utils.py
import customtkinter as ctk
import tkinter.messagebox as messagebox
import sys
from gui_utils import SafeCTk, bind_entry_shortcuts, create_context_menu # Используем SafeCTk и другие утилиты

APP_PASSWORD = "qwer" # Держите пароль здесь или загружайте из конфигурации

def ask_password(on_success_callback):
    """
    Отображает окно запроса пароля.

    Args:
        on_success_callback: Функция, которая будет вызвана после успешного ввода пароля.
    """
    password_window = None # Объявляем заранее

    def check_password():
        nonlocal password_window # Для доступа к внешней переменной
        if entry.get() == APP_PASSWORD:
            if password_window:
                password_window.destroy() # Закрываем окно пароля
            on_success_callback()
        else:
            messagebox.showerror("Помилка", "Невірний пароль!")
            entry.delete(0, ctk.END) # Очищаем поле ввода

    def on_close_password_window():
        nonlocal password_window
        if password_window:
            password_window.destroy()
        sys.exit(0) # Завершаем приложение, если окно пароля закрыто

    try:
        password_window = SafeCTk() # Используем SafeCTk
        password_window.title("Введіть пароль")
        password_window.geometry("300x150") # Немного увеличим высоту для отступов
        password_window.resizable(False, False)
        password_window.protocol("WM_DELETE_WINDOW", on_close_password_window)

        # Чтобы окно было поверх других и в центре
        password_window.attributes('-topmost', True)
        password_window.lift()
        password_window.after(10, lambda: password_window.attributes('-topmost', False)) # Убираем topmost после отображения
        password_window.eval('tk::PlaceWindow . center')


        main_frame = ctk.CTkFrame(password_window)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        ctk.CTkLabel(main_frame, text="Введіть 4-значний пароль:").pack(pady=(0,10))
        entry = ctk.CTkEntry(main_frame, show="*", justify="center", font=("Arial", 18), width=150)
        entry.pack(pady=5)
        entry.focus()

        # Контекстное меню для поля ввода пароля
        # Создаем временное root окно для контекстного меню, если оно еще не создано.
        # В идеале, root должен передаваться, но для простоты так:
        temp_root_for_menu = password_window # Используем окно пароля как временный root
        password_context_menu = create_context_menu(temp_root_for_menu)
        bind_entry_shortcuts(entry, password_context_menu) # Передаем объект меню

        ctk.CTkButton(main_frame, text="Увійти", command=check_password, width=100).pack(pady=10)

        # Обработка нажатия Enter
        entry.bind("<Return>", lambda event: check_password())

        password_window.mainloop()

    except Exception as e:
        # Используем ваш логгер, если он будет импортирован
        print(f"Ошибка при создании окна пароля: {str(e)}\n{traceback.format_exc()}")
        # from error_handler import log_and_show_error
        # log_and_show_error(type(e), e, e.__traceback__)
        # messagebox.showerror("Критична помилка", "Не вдалося відкрити вікно пароля.")
        sys.exit(1)