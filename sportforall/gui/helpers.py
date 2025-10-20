# sportforall/gui/helpers.py

import tkinter as tk
from tkinter import ttk, scrolledtext
import sys

# --- Допоміжний клас для скрольованого фрейму ---
class ScrolledFrame(ttk.Frame):
    """
    Фрейм з вертикальним скролбаром.
    Розміщуйте вміст всередині 'self.inner_frame'.
    """
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        self.canvas = tk.Canvas(self)
        self.vscrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vscrollbar.set)

        self.vscrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Створюємо внутрішній фрейм, який буде скролитись
        self.inner_frame = ttk.Frame(self.canvas)
        # Розміщуємо внутрішній фрейм у вікні канвасу
        self._canvas_window_id = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        # Прив'язуємо події зміни розміру фрейму та канвасу для оновлення області прокрутки
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Підтримка скролінгу колесом миші
        # Прив'язуємо до кореневого вікна, щоб скролінг працював, коли курсор над будь-яким віджетом
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        # Для Linux (кнопки 4 і 5 миші)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel) # Scroll up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel) # Scroll down


    def _on_frame_configure(self, event=None):
        """Оновлює область прокрутки канвасу при зміні розміру внутрішнього фрейму."""
        # Додаємо невеликий відступ знизу, щоб останній елемент не був перекритий
        bbox = self.canvas.bbox("all")
        if bbox:
             x1, y1, x2, y2 = bbox
             # Встановлюємо scrollregion з урахуванням розміру внутрішнього фрейму
             self.canvas.config(scrollregion=bbox)
             # Додаємо невеликий додатковий відступ для зручності прокрутки до кінця
             self.canvas.config(scrollregion=(x1, y1, x2, y2 + 10))


    def _on_canvas_configure(self, event=None):
        """Оновлює ширину внутрішнього фрейму при зміні ширини канвасу."""
        canvas_width = event.width
        # Налаштовуємо ширину вікна в канвасі, щоб внутрішній фрейм розтягувався
        self.canvas.itemconfigure(self._canvas_window_id, width=canvas_width)


    def _on_mousewheel(self, event):
        """Обробляє скролінг колесом миші."""
        # Визначаємо віджет під курсором миші
        widget_at_mouse = event.widget.winfo_containing(event.x_root, event.y_root)
        if widget_at_mouse is not None:
             # Перевіряємо, чи знаходиться віджет під курсором всередині цього ScrolledFrame (або його inner_frame/canvas)
             is_within_frame = False
             current = widget_at_mouse
             while current is not None:
                  # Якщо знайшли сам канвас або inner_frame, значить курсор над нашою областю
                  if current == self.canvas or current == self.inner_frame:
                       is_within_frame = True
                       break
                  # Піднімаємось по ієрархії батьківських віджетів
                  current = current.master

             # Якщо курсор над нашим скролбаром, не скролимо канвас
             if current == self.vscrollbar:
                  is_within_frame = False


             if is_within_frame:
                # Визначаємо напрямок та величину скролінгу залежно від ОС
                if sys.platform.startswith('win'):
                    # event.delta зазвичай 120 для однієї "одиниці" скролінгу
                    # Множимо на -1, бо WheelUp має скролити вгору (зменшення y)
                    self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                elif sys.platform.startswith('linux'):
                    # Для Linux Button 4 - вгору, Button 5 - вниз
                    if event.num == 4: # Прокрутка вгору
                        self.canvas.yview_scroll(-1, "units")
                    elif event.num == 5: # Прокрутка вниз
                        self.canvas.yview_scroll(1, "units")


# --- Контекстне меню для полів введення ---
class TextEntryMenu(tk.Menu):
    """
    Контекстне меню (ПКМ) для стандартних полів введення тексту (Entry, Text, ScrolledText).
    """
    def __init__(self, master=None, **kwargs):
        super().__init__(master, tearoff=0, **kwargs)
        # Використовуємо standard events для Cut/Copy/Paste
        self.add_command(label="Вирізати", command=self.cut_text)
        self.add_command(label="Копіювати", command=self.copy_text)
        self.add_command(label="Вставити", command=self.paste_text)
        self.add_separator()
        self.add_command(label="Виділити все", command=self.select_all_text)

        self.active_widget = None # Віджет, на якому було викликано ПКМ

    def show(self, event):
        # Визначаємо віджет під курсором миші, де викликано ПКМ
        widget_at_mouse = event.widget.winfo_containing(event.x_root, event.y_root)
        # Перевіряємо, чи є цей віджет полем введення тексту
        if widget_at_mouse is not None and isinstance(widget_at_mouse, (ttk.Entry, tk.Entry, tk.Text, scrolledtext.ScrolledText)):
            self.active_widget = widget_at_mouse
            try:
                 # Налаштовуємо стан пунктів меню (Вирізати/Вставити активні, якщо поле не read-only)
                 state_cut_paste = "normal"
                 # tk.TclError може виникнути, якщо віджет вже видалено або не валідний
                 try:
                     if self.active_widget.cget('state') == 'readonly':
                         state_cut_paste = "disabled"
                 except tk.TclError:
                     state_cut_paste = "disabled" # Якщо виникла помилка, вважаємо віджет неактивним


                 self.entryconfig("Вирізати", state=state_cut_paste)
                 self.entryconfig("Вставити", state=state_cut_paste)
                 self.entryconfig("Копіювати", state="normal") # Копіювати завжди активно
                 self.entryconfig("Виділити все", state="normal") # Виділити все завжди активно

                 # Показуємо меню під курсором миші
                 self.tk_popup(event.x_root, event.y_root)
            finally:
                # Після показу меню знімаємо "захват" миші
                self.grab_release()
        else:
             # Якщо ПКМ викликано не на текстовому полі, не показуємо меню і скидаємо active_widget
             self.active_widget = None


    def cut_text(self):
        if self.active_widget:
             try:
                 # Генеруємо стандартну подію "Cut" для віджета
                 self.active_widget.event_generate("<<Cut>>")
             except tk.TclError:
                 # Ігноруємо помилки, якщо віджет вже не валідний
                 pass

    def copy_text(self):
         if self.active_widget:
             try:
                 # Генеруємо стандартну подію "Copy" для віджета
                 self.active_widget.event_generate("<<Copy>>")
             except tk.TclError:
                 pass

    def paste_text(self):
        if self.active_widget:
            try:
                 # Генеруємо стандартну подію "Paste" для віджета
                 self.active_widget.event_generate("<<Paste>>")
            except tk.TclError:
                 pass


    def select_all_text(self):
        if self.active_widget:
            # Виділяємо весь текст у віджеті
            try:
                if isinstance(self.active_widget, (ttk.Entry, tk.Entry)):
                     self.active_widget.selection_range(0, tk.END)
                     # Переміщаємо курсор в кінець
                     self.active_widget.icursor(tk.END)
                elif isinstance(self.active_widget, (tk.Text, scrolledtext.ScrolledText)):
                     # Для Text та ScrolledText використовуємо теги
                     self.active_widget.tag_add(tk.SEL, "1.0", tk.END)
                     # Переміщаємо точку вставки (курсор) в кінець
                     self.active_widget.mark_set(tk.INSERT, tk.END)
                # Передаємо фокус на віджет, щоб виділення було видно
                self.active_widget.focus_set()
            except tk.TclError:
                 pass