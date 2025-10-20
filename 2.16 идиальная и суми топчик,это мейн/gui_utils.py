# gui_utils.py
import customtkinter as ctk
import tkinter as tk

# Глобальные переменные для доступа из функций
context_menu_tk_popup = None # Для хранения tk_popup из контекстного меню

# Словарь для хранения активных after ID
after_callbacks = {}

def safe_after_cancel(widget, after_id):
    """Безопасная отмена отложенного вызова"""
    try:
        if widget and after_id:
            widget.after_cancel(after_id)
    except Exception:
        pass  # Игнорировать ошибки при отмене

def cleanup_after_callbacks():
    """Очистка всех отложенных вызовов"""
    global after_callbacks
    for widget_id in list(after_callbacks.keys()):
        widget = None
        try: # Пытаемся получить виджет по ID, если он еще существует
            widget = ctk.CTk._get_widget_by_id(int(widget_id)) # CTk._get_widget_by_id ожидает int
        except: # Если виджет не найден (уже уничтожен), пропускаем
             pass

        if widget:
            for after_id in after_callbacks.get(widget_id, []):
                safe_after_cancel(widget, after_id)
    after_callbacks.clear()


def safe_after(widget, ms, callback):
    """Безопасный вызов after с отслеживанием идентификаторов"""
    if not widget:
        return None

    try:
        # Используем id(widget) как более надежный идентификатор, чем str(widget)
        # который может меняться или быть не уникальным в некоторых сценариях Tkinter.
        widget_id = str(id(widget))
        after_id = widget.after(ms, callback)

        if widget_id not in after_callbacks:
            after_callbacks[widget_id] = []

        after_callbacks[widget_id].append(after_id)
        return after_id
    except Exception:
        return None

# Переопределяем методы CustomTkinter
original_after_ctk = ctk.CTk.after
def patched_after_ctk(self, ms, func=None, *args):
    """Патч метода after для CTk для отслеживания вызовов"""
    if func:
        callback = lambda: func(*args) if args else func()
        return safe_after(self, ms, callback)
    else:
        # Для вызовов вроде `id = widget.after(ms)` без функции
        return original_after_ctk(self, ms)


ctk.CTk.after = patched_after_ctk

# Патч для CTkBaseClass
original_base_after_ctk = ctk.CTkBaseClass.after
def patched_base_after_ctk(self, ms, func=None, *args):
    """Патч метода after для CTkBaseClass для отслеживания вызовов"""
    if func:
        callback = lambda: func(*args) if args else func()
        return safe_after(self, ms, callback)
    else:
        return original_base_after_ctk(self, ms)

ctk.CTkBaseClass.after = patched_base_after_ctk


# Пользовательский класс окна CustomTkinter с защитой от after-ошибок
class SafeCTk(ctk.CTk):
    def destroy(self):
        cleanup_after_callbacks() # Очищаем колбэки перед уничтожением окна
        super().destroy()


def bind_entry_shortcuts(entry_widget, context_menu_ref):
    """Привязывает стандартные сочетания клавиш и контекстное меню к CTkEntry."""
    try:
        # Убедимся, что context_menu_ref это действительно объект Menu
        if not isinstance(context_menu_ref, tk.Menu):
            # print("Warning: context_menu_ref is not a tk.Menu object in bind_entry_shortcuts")
            return

        internal_entry = entry_widget._entry # Доступ к внутреннему tk.Entry

        def select_all(event):
            internal_entry.select_range(0, "end")
            return "break"

        def copy_text(event):
            if internal_entry.selection_present():
                internal_entry.event_generate('<<Copy>>')
            return "break"

        def paste_text(event):
            internal_entry.event_generate('<<Paste>>')
            return "break"

        def cut_text(event):
            if internal_entry.selection_present():
                internal_entry.event_generate('<<Cut>>')
            return "break"

        def show_context_menu(event):
            # Сохраняем текущий виджет в фокусе для команд меню
            context_menu_ref.current_widget = internal_entry
            context_menu_ref.tk_popup(event.x_root, event.y_root)
            return "break"

        # Привязываем к внутреннему tk.Entry
        internal_entry.bind("<Control-a>", select_all)
        internal_entry.bind("<Control-A>", select_all) # Для верхнего регистра
        internal_entry.bind("<Button-3>", show_context_menu)
        internal_entry.bind("<Control-c>", copy_text)
        internal_entry.bind("<Control-C>", copy_text)
        internal_entry.bind("<Control-v>", paste_text)
        internal_entry.bind("<Control-V>", paste_text)
        internal_entry.bind("<Control-x>", cut_text)
        internal_entry.bind("<Control-X>", cut_text)

        # Также привяжем к самому CTkEntry, на всякий случай, если CTk перехватывает что-то
        entry_widget.bind("<Control-a>", select_all, add="+")
        entry_widget.bind("<Control-A>", select_all, add="+")
        entry_widget.bind("<Button-3>", show_context_menu, add="+") # Это может конфликтовать, но посмотрим
        entry_widget.bind("<Control-c>", copy_text, add="+")
        entry_widget.bind("<Control-C>", copy_text, add="+")
        entry_widget.bind("<Control-v>", paste_text, add="+")
        entry_widget.bind("<Control-V>", paste_text, add="+")
        entry_widget.bind("<Control-x>", cut_text, add="+")
        entry_widget.bind("<Control-X>", cut_text, add="+")


    except Exception as e:
        print(f"Ошибка в bind_entry_shortcuts: {e}")
        # Используйте ваш логгер ошибок, если он доступен здесь
        # from error_handler import log_and_show_error
        # log_and_show_error(type(e), e, e.__traceback__)

def create_context_menu(root_window):
    """Создает и возвращает контекстное меню."""
    context_menu = tk.Menu(root_window, tearoff=0)

    def get_focused_widget():
        # Пытаемся получить виджет из свойства, установленного в show_context_menu
        widget = getattr(context_menu, 'current_widget', root_window.focus_get())
        # Проверяем, является ли это внутренним tk.Entry от CTkEntry
        if hasattr(widget, 'master') and isinstance(widget.master, ctk.CTkEntry):
             return widget # Это внутренний tk.Entry
        # Если это сам CTkEntry, пытаемся получить его внутренний _entry
        if isinstance(widget, ctk.CTkEntry):
            return widget._entry
        return widget


    context_menu.add_command(label="Копіювати", command=lambda: get_focused_widget().event_generate('<<Copy>>'))
    context_menu.add_command(label="Вставити", command=lambda: get_focused_widget().event_generate('<<Paste>>'))
    context_menu.add_command(label="Вирізати", command=lambda: get_focused_widget().event_generate('<<Cut>>'))
    context_menu.add_separator()
    context_menu.add_command(label="Виділити все", command=lambda: get_focused_widget().select_range(0, 'end'))
    return context_menu