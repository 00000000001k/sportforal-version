# gui_utils.py
import customtkinter as ctk
import tkinter as tk
import pyperclip  # Для надійної роботи з буфером обміну
import threading

# Глобальні змінні для доступу з функцій
context_menu_tk_popup = None
after_callbacks = {}
entry_widgets_list = []  # Список для навігації стрілками


def safe_after_cancel(widget, after_id):
    """Безпечна відміна відкладеного виклику"""
    try:
        if widget and after_id:
            widget.after_cancel(after_id)
    except Exception:
        pass


def cleanup_after_callbacks():
    """Очищення всіх відкладених викликів"""
    global after_callbacks
    for widget_id in list(after_callbacks.keys()):
        widget = None
        try:
            widget = ctk.CTk._get_widget_by_id(int(widget_id))
        except:
            pass

        if widget:
            for after_id in after_callbacks.get(widget_id, []):
                safe_after_cancel(widget, after_id)
    after_callbacks.clear()


def safe_after(widget, ms, callback):
    """Безпечний виклик after з відстеженням ідентифікаторів"""
    if not widget:
        return None

    try:
        widget_id = str(id(widget))
        after_id = widget.after(ms, callback)

        if widget_id not in after_callbacks:
            after_callbacks[widget_id] = []

        after_callbacks[widget_id].append(after_id)
        return after_id
    except Exception:
        return None


# Патчі для CustomTkinter
original_after_ctk = ctk.CTk.after


def patched_after_ctk(self, ms, func=None, *args):
    """Патч методу after для CTk"""
    if func:
        callback = lambda: func(*args) if args else func()
        return safe_after(self, ms, callback)
    else:
        return original_after_ctk(self, ms)


ctk.CTk.after = patched_after_ctk

original_base_after_ctk = ctk.CTkBaseClass.after


def patched_base_after_ctk(self, ms, func=None, *args):
    """Патч методу after для CTkBaseClass"""
    if func:
        callback = lambda: func(*args) if args else func()
        return safe_after(self, ms, callback)
    else:
        return original_base_after_ctk(self, ms)


ctk.CTkBaseClass.after = patched_base_after_ctk


class SafeCTk(ctk.CTk):
    def destroy(self):
        cleanup_after_callbacks()
        super().destroy()


def safe_copy_to_clipboard(text):
    """Безпечне копіювання в буфер обміну"""
    try:
        # Спочатку пробуємо pyperclip
        pyperclip.copy(text)
    except Exception:
        try:
            # Якщо pyperclip не працює, використовуємо tkinter
            root = tk._default_root
            if root:
                root.clipboard_clear()
                root.clipboard_append(text)
                root.update()
        except Exception:
            pass


def safe_paste_from_clipboard():
    """Безпечне отримання з буферу обміну"""
    try:
        # Спочатку пробуємо pyperclip
        return pyperclip.paste()
    except Exception:
        try:
            # Якщо pyperclip не працює, використовуємо tkinter
            root = tk._default_root
            if root:
                return root.clipboard_get()
        except Exception:
            return ""


def register_entry_for_navigation(entry_widget):
    """Реєструє віджет для навігації стрілками"""
    global entry_widgets_list
    if entry_widget not in entry_widgets_list:
        entry_widgets_list.append(entry_widget)


def navigate_entries(current_entry, direction):
    """Навігація між полями стрілками вгору/вниз"""
    global entry_widgets_list

    try:
        # Фільтруємо тільки видимі та активні віджети
        active_entries = [e for e in entry_widgets_list if e.winfo_exists() and e.winfo_viewable()]

        if current_entry not in active_entries:
            return

        current_index = active_entries.index(current_entry)

        if direction == "up" and current_index > 0:
            active_entries[current_index - 1].focus_set()
        elif direction == "down" and current_index < len(active_entries) - 1:
            active_entries[current_index + 1].focus_set()

    except (ValueError, IndexError, tk.TclError):
        pass


def bind_entry_shortcuts(entry_widget, context_menu_ref):
    """Прив'язує стандартні комбінації клавіш та контекстне меню до CTkEntry"""
    try:
        if not isinstance(context_menu_ref, tk.Menu):
            return

        # Реєструємо віджет для навігації
        register_entry_for_navigation(entry_widget)

        internal_entry = entry_widget._entry

        def select_all(event):
            """Виділити все"""
            try:
                internal_entry.select_range(0, "end")
                internal_entry.icursor("end")
            except Exception:
                pass
            return "break"

        def copy_text(event):
            """Копіювати текст"""
            try:
                if internal_entry.selection_present():
                    selected_text = internal_entry.selection_get()
                    safe_copy_to_clipboard(selected_text)
                else:
                    # Якщо нічого не виділено, копіюємо весь текст
                    all_text = internal_entry.get()
                    safe_copy_to_clipboard(all_text)
            except Exception:
                pass
            return "break"

        def paste_text(event):
            """Вставити текст"""
            try:
                clipboard_text = safe_paste_from_clipboard()
                if clipboard_text:
                    # Видаляємо виділений текст, якщо є
                    if internal_entry.selection_present():
                        internal_entry.delete("sel.first", "sel.last")
                    # Вставляємо текст в позицію курсору
                    insert_pos = internal_entry.index("insert")
                    internal_entry.insert(insert_pos, clipboard_text)
            except Exception:
                pass
            return "break"

        def cut_text(event):
            """Вирізати текст"""
            try:
                if internal_entry.selection_present():
                    selected_text = internal_entry.selection_get()
                    safe_copy_to_clipboard(selected_text)
                    internal_entry.delete("sel.first", "sel.last")
            except Exception:
                pass
            return "break"

        def navigate_up(event):
            """Навігація вгору"""
            navigate_entries(entry_widget, "up")
            return "break"

        def navigate_down(event):
            """Навігація вниз"""
            navigate_entries(entry_widget, "down")
            return "break"

        def show_context_menu(event):
            """Показати контекстне меню"""
            try:
                context_menu_ref.current_widget = internal_entry
                context_menu_ref.tk_popup(event.x_root, event.y_root)
            except Exception:
                pass
            return "break"

        # Прив'язуємо до внутрішнього tk.Entry
        bindings = [
            ("<Control-a>", select_all),
            ("<Control-A>", select_all),
            ("<Control-c>", copy_text),
            ("<Control-C>", copy_text),
            ("<Control-v>", paste_text),
            ("<Control-V>", paste_text),
            ("<Control-x>", cut_text),
            ("<Control-X>", cut_text),
            ("<Button-3>", show_context_menu),
            ("<Up>", navigate_up),
            ("<Down>", navigate_down),
        ]

        for key_combo, handler in bindings:
            try:
                internal_entry.bind(key_combo, handler)
                entry_widget.bind(key_combo, handler, add="+")
            except Exception:
                pass

    except Exception as e:
        print(f"Помилка в bind_entry_shortcuts: {e}")


def create_context_menu(root_window):
    """Створює та повертає контекстне меню"""
    context_menu = tk.Menu(root_window, tearoff=0)

    def get_focused_widget():
        widget = getattr(context_menu, 'current_widget', root_window.focus_get())
        if hasattr(widget, 'master') and isinstance(widget.master, ctk.CTkEntry):
            return widget
        if isinstance(widget, ctk.CTkEntry):
            return widget._entry
        return widget

    def context_copy():
        """Копіювати через контекстне меню"""
        try:
            widget = get_focused_widget()
            if hasattr(widget, 'selection_present') and widget.selection_present():
                text = widget.selection_get()
            else:
                text = widget.get()
            safe_copy_to_clipboard(text)
        except Exception:
            pass

    def context_paste():
        """Вставити через контекстне меню"""
        try:
            widget = get_focused_widget()
            clipboard_text = safe_paste_from_clipboard()
            if clipboard_text:
                if hasattr(widget, 'selection_present') and widget.selection_present():
                    widget.delete("sel.first", "sel.last")
                insert_pos = widget.index("insert") if hasattr(widget, 'index') else 0
                widget.insert(insert_pos, clipboard_text)
        except Exception:
            pass

    def context_cut():
        """Вирізати через контекстне меню"""
        try:
            widget = get_focused_widget()
            if hasattr(widget, 'selection_present') and widget.selection_present():
                text = widget.selection_get()
                safe_copy_to_clipboard(text)
                widget.delete("sel.first", "sel.last")
        except Exception:
            pass

    def context_select_all():
        """Виділити все через контекстне меню"""
        try:
            widget = get_focused_widget()
            if hasattr(widget, 'select_range'):
                widget.select_range(0, 'end')
        except Exception:
            pass

    context_menu.add_command(label="Копіювати", command=context_copy)
    context_menu.add_command(label="Вставити", command=context_paste)
    context_menu.add_command(label="Вирізати", command=context_cut)
    context_menu.add_separator()
    context_menu.add_command(label="Виділити все", command=context_select_all)

    return context_menu