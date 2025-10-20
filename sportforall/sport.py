# sport.py
import customtkinter as ctk
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
# import json # Уже не нужен здесь напрямую, используется в data_persistence
import os
# import re # Уже не нужен здесь напрямую, используется в custom_widgets
# from openpyxl import Workbook, load_workbook # Workbook здесь не нужен, load_workbook тоже
import traceback  # Нужен для явных блоков try-except, если такие останутся
import datetime  # Нужен для явных блоков try-except, если такие останутся
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
    # input("Натисніть Enter для завершення...") # log_and_show_error уже ждет
    sys.exit(1)

# Глобальные переменные для основного приложения
FIELDS = [
    "товар", "дк", "захід", "дата", "адреса", "сума",
    "сума прописом", "пдв", "кількість", "ціна за одиницю",
    # "загальна сума", # Это поле, кажется, дублирует "разом" или "сума" и заполняется в кошторисе. Уточнить.
    "разом"  # "разом" обычно это сумма по одной позиции, а "загальна сума" - итог.
    # В вашем EXAMPLES "загальна сума" и "разом" одинаковы.
    # Если "загальна сума" это итоговая по всем товарам, то она не должна быть здесь как поле для каждого договора.
]

EXAMPLES = {
    "товар": "наприклад: медалі зі стрічкою",
    "дк": "наприклад: ДК 021:2015: 18512200-3",
    "захід": "наприклад: 4 етапу “Фізкультурно-оздоровчих заходів...",
    "дата": "наприклад: з 06 по 09 травня, з 13 по 16 травня 2025 року",
    "адреса": "наприклад: КП МСК “Дніпро”, вул. Смілянська, 78, м. Черкаси.",
    "сума": "наприклад: 15120.00",  # Используем точки для ввода, CustomEntry обработает
    "сума прописом": "Автоматично (П’ятнадцять тисяч сто двадцять грн 00 коп.)",
    "пдв": "наприклад: без ПДВ",
    "кількість": "наприклад: 144",
    "ціна за одиницю": "наприклад: 105.00",  # Изменил, чтобы сумма была 15120
    "разом": "Автоматично (15120.00)"  # Это поле тоже можно сделать автоматическим
}

# Глобальные переменные для GUI основного приложения
document_blocks = []  # Список для хранения информации о блоках документов
main_app_root = None
scroll_frame_main = None
main_context_menu = None


def make_show_hint_command(hint_text, field_name):
    """Создает команду для кнопки подсказки, отображающую messagebox."""

    def show_hint():
        messagebox.showinfo(f"Підказка для <{field_name}>", hint_text)

    return show_hint


def generate_documents_word():
    """Генерирует документы Word на основе заполненных данных."""
    global document_blocks  # Убедимся, что используем глобальную переменную
    if not document_blocks:
        messagebox.showwarning("Увага", "Не додано жодного договору для генерації документів Word.")
        return False

    # Проверка на заполненность обязательных полей (если необходимо)
    # for i, block in enumerate(document_blocks, start=1):
    #     # ... (ваша логика проверки) ...

    save_dir = filedialog.askdirectory(title="Оберіть папку для збереження документів Word")
    if not save_dir:
        return False

    # Сохраняем данные для каждого блока по отдельности в JSON
    for block in document_blocks:
        if "path" in block and block["path"]:  # Убедимся, что путь к шаблону есть
            block_data = {f: block["entries"][f].get() for f in FIELDS if f in block["entries"]}
            save_memory(block_data, block["path"])  # Сохраняем по ключу шаблона

    word_app = None
    try:
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False  # Работа в фоновом режиме
        generated_count = 0

        for block in document_blocks:
            template_path_abs = os.path.abspath(block["path"])
            if not os.path.exists(template_path_abs):
                log_and_show_error(FileNotFoundError, f"Шаблон не знайдено: {template_path_abs}", None)
                continue

            try:
                doc = word_app.Documents.Open(template_path_abs)
                # Замена плейсхолдеров
                for key in FIELDS:
                    if key in block["entries"]:
                        # Используем Range.Find.Execute для более надежной замены
                        find_obj = doc.Content.Find
                        find_obj.ClearFormatting()
                        find_obj.Replacement.ClearFormatting()
                        find_obj.Execute(FindText=f"<{key}>",
                                         MatchCase=False,
                                         MatchWholeWord=False,
                                         MatchWildcards=False,
                                         MatchSoundsLike=False,
                                         MatchAllWordForms=False,
                                         Forward=True,
                                         Wrap=win32.constants.wdFindContinue,  # Константа для продолжения поиска
                                         Format=False,
                                         ReplaceWith=block["entries"][key].get(),
                                         Replace=win32.constants.wdReplaceAll)  # Константа для замены всех вхождений

                base_name = os.path.splitext(os.path.basename(block["path"]).replace("ШАБЛОН", "").strip())[0]
                # Формируем имя файла, избегая недопустимых символов из поля "товар"
                товар_name = block['entries']['товар'].get()
                safe_товар_name = "".join(c if c.isalnum() or c in " -" else "_" for c in
                                          товар_name)  # Оставляем буквы, цифры, пробелы, дефисы
                safe_товар_name = safe_товар_name[:50]  # Ограничиваем длину

                output_filename_word = f"{base_name} {safe_товар_name}.docm"  # или .docx, если макросы не нужны
                save_path_word = os.path.join(save_dir, output_filename_word)

                doc.SaveAs(save_path_word, FileFormat=13)  # 13 для .docm (Word Macro-Enabled Document)
                # 16 для .docx (Word Document)
                doc.Close(False)  # False - не сохранять изменения при закрытии (уже сохранили)
                generated_count += 1
            except Exception as e_doc:
                # Логируем ошибку для конкретного документа, но продолжаем для остальных
                log_and_show_error(type(e_doc), f"Помилка при обробці шаблону: {block['path']}\n{e_doc}",
                                   sys.exc_info()[2])
                if 'doc' in locals() and doc is not None:  # Если документ был открыт, пытаемся его закрыть
                    try:
                        doc.Close(False)
                    except:
                        pass

        if generated_count > 0:
            messagebox.showinfo("Успіх", f"{generated_count} документ(и) Word збережено успішно в папці:\n{save_dir}")
        elif not document_blocks:
            pass  # Уже было сообщение
        else:
            messagebox.showwarning("Увага", "Жодного документа Word не було згенеровано через помилки.")
        return True

    except AttributeError as e_attr:  # Часто возникает, если COM объект неправильно инициализирован или Word закрыт
        log_and_show_error(type(e_attr),
                           f"Помилка доступу до Word (можливо, Word не встановлено або проблема з COM): {e_attr}",
                           sys.exc_info()[2])
        # messagebox.showerror("Помилка MS Word", f"Не вдалося взаємодіяти з MS Word: {e_attr}\nПереконайтеся, що MS Word встановлено та налаштовано.")
        return False
    except Exception as e_main_word:
        log_and_show_error(type(e_main_word), f"Загальна помилка при генерації документів Word: {e_main_word}",
                           sys.exc_info()[2])
        return False
    finally:
        if word_app:
            try:
                word_app.Quit()
            except:
                pass  # Игнорируем ошибки при выходе из Word


def combined_generation_process():
    """Объединенный процесс генерации: сначала Excel, потом Word, потом Кошторис."""
    global document_blocks
    if not document_blocks:
        messagebox.showwarning("Увага", "Не додано жодного договору для генерації.")
        return

    # 1. Проверка на заполненность полей (можно вынести в отдельную функцию)
    for i, block in enumerate(document_blocks, start=1):
        for field in FIELDS:  # Используем глобальные FIELDS
            entry_widget = block["entries"].get(field)
            # Поле "сума прописом" и "разом" может быть автоматическим, его не проверяем на пустоту если оно readonly
            if field in ["сума прописом", "разом"] and entry_widget and entry_widget.cget("state") == "readonly":
                continue
            if not entry_widget or not entry_widget.get().strip():
                messagebox.showerror("Помилка заповнення", f"Блок договору #{i}: поле <{field}> порожнє.")
                return

    # 2. Генерация Excel с данными договоров
    if not export_document_data_to_excel(document_blocks, FIELDS):  # Передаем FIELDS
        messagebox.showerror("Помилка",
                             "Не вдалося згенерувати Excel файл з даними договорів. Подальша генерація скасована.")
        return

    # 3. Генерация документов Word
    if not generate_documents_word():  # generate_documents_word сама покажет сообщение об успехе/ошибке
        messagebox.showwarning("Увага", "Генерація документів Word не була повністю успішною або скасована.")
        # Решите, продолжать ли с кошторисом, если Word не удался. Пока продолжаем.

    # 4. Заполнение кошториса
    if koshtorys.fill_koshtorys(document_blocks):  # fill_koshtorys сама покажет сообщение
        messagebox.showinfo("Завершено", "Усі документи (Excel, Word, Кошторис) оброблено.")
    else:
        messagebox.showerror("Помилка Кошторису", "Не вдалося заповнити файл кошторису.")


# ---------------- ОСНОВНИЙ ІНТЕРФЕЙС ----------------
def launch_main_app():
    global main_app_root, scroll_frame_main, document_blocks, main_context_menu, FIELDS, EXAMPLES

    try:
        ctk.set_appearance_mode("light")  # или "dark", "system"
        ctk.set_default_color_theme("green")  # или "green", "dark-blue"

        main_app_root = SafeCTk()  # Используем наш SafeCTk
        main_app_root.title("SportForAll  v2.8+")  # Обновим версию
        main_app_root.geometry("1200x750")  # Немного больше места

        def on_root_close_main_app():
            if messagebox.askokcancel("Вихід", "Ви впевнені, що хочете вийти?"):
                # gui_utils.cleanup_after_callbacks() # SafeCTk.destroy() уже это делает
                main_app_root.destroy()
                # sys.exit(0) # Не нужно, destroy и так завершит главный цикл Tkinter

        main_app_root.protocol("WM_DELETE_WINDOW", on_root_close_main_app)

        # --- Контекстное меню ---
        main_context_menu = create_context_menu(main_app_root)  # Используем функцию из gui_utils

        # --- Верхняя панель с кнопками ---
        top_controls_frame = ctk.CTkFrame(main_app_root)
        top_controls_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkButton(top_controls_frame, text="➕ Додати договір", command=lambda: add_new_template_block(),
                      fg_color="#2196F3").pack(side="left", padx=5)
        ctk.CTkButton(top_controls_frame, text="📄 Згенерувати ВСІ документи", command=combined_generation_process,
                      fg_color="#4CAF50").pack(side="left", padx=5)
        ctk.CTkButton(top_controls_frame, text="💰 Тільки Кошторис",
                      command=lambda: koshtorys.fill_koshtorys(document_blocks),
                      fg_color="#FF9800").pack(side="left", padx=5)
        # Кнопка экспорта в Excel (дополнительно к общей генерации)
        ctk.CTkButton(top_controls_frame, text="📥 Експорт в Excel",
                      command=lambda: export_document_data_to_excel(document_blocks, FIELDS),
                      fg_color="#00BCD4").pack(side="left", padx=5)

        version_label = ctk.CTkLabel(top_controls_frame, text="version 2.8+", text_color="gray", font=("Arial", 12))
        version_label.pack(side="right", padx=10)

        # --- Скроллируемый контейнер для блоков договоров ---
        scroll_frame_main = ctk.CTkScrollableFrame(main_app_root, width=1100, height=600)  # Имя изменено
        scroll_frame_main.pack(padx=10, pady=(0, 10), fill="both", expand=True)

        def add_new_template_block():
            """Запрашивает шаблон и добавляет блок полей для него."""
            filepath = filedialog.askopenfilename(
                title="Оберіть шаблон договору (.docm)",
                filetypes=[("Word Macro-Enabled Documents", "*.docm"), ("Word Documents", "*.docx"),
                           ("All files", "*.*")]
            )
            if filepath:
                create_document_fields_block(filepath)

        def create_document_fields_block(template_filepath):
            """Создает GUI блок с полями для одного шаблона договора."""
            global document_blocks  # document_blocks изменяется (append), поэтому объявляем global

            # # Проверяем, не добавлен ли уже такой шаблон (по пути)
            # for existing_block in document_blocks:
            #     if existing_block["path"] == template_filepath:
            #         messagebox.showwarning("Увага", f"Шаблон '{os.path.basename(template_filepath)}' вже додано.")
            #         return

            block_frame = ctk.CTkFrame(scroll_frame_main, border_width=1, border_color="gray70")
            block_frame.pack(pady=10, padx=5, fill="x")

            # --- Заголовок блока (путь к шаблону и кнопка удаления) ---
            header_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
            header_frame.pack(fill="x", padx=5, pady=(5, 0))

            # Используем CTkLabel для пути, чтобы он мог быть длинным и переноситься (если настроить)
            path_label = ctk.CTkLabel(header_frame,
                                      text=f"Шаблон: {os.path.basename(template_filepath)} ({template_filepath})",
                                      text_color="blue", anchor="w", wraplength=800)  # wraplength для переноса
            path_label.pack(side="left", padx=(0, 5), expand=True, fill="x")

            current_block_entries = {}  # Словарь для хранения виджетов Entry этого блока

            # Загружаем данные из памяти для этого конкретного шаблона
            template_specific_memory = get_template_memory(template_filepath)
            # Загружаем общие данные из памяти (для обратной совместимости или как fallback)
            general_memory = load_memory()  # Это вернет весь JSON, нужно будет извлекать по ключам полей

            # --- Поля ввода ---
            fields_grid_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
            fields_grid_frame.pack(fill="x", padx=5, pady=5)
            # fields_grid_frame.columnconfigure(1, weight=1) # Поле ввода будет расширяться

            for i, field_key in enumerate(FIELDS):
                # fields_grid_frame.rowconfigure(i, weight=1) # Равномерное распределение по высоте если нужно

                label = ctk.CTkLabel(fields_grid_frame, text=f"<{field_key}>", anchor="w", width=140,
                                     font=("Arial", 12))  # Немного шире
                label.grid(row=i, column=0, padx=5, pady=3, sticky="w")

                entry = CustomEntry(fields_grid_frame, field_name=field_key,
                                    examples_dict=EXAMPLES)  # width убран, будет через sticky="ew"
                entry.grid(row=i, column=1, padx=5, pady=3, sticky="ew")
                fields_grid_frame.columnconfigure(1, weight=1)  # Колонка с Entry будет растягиваться

                # Вставка сохраненных данных: сначала из специфичных для шаблона, потом из общих
                saved_value = template_specific_memory.get(field_key, general_memory.get(field_key))
                if saved_value is not None:  # Проверяем на None, так как пустая строка тоже валидное сохраненное значение
                    entry.set_text(saved_value)  # Используем set_text для корректной установки и обработки плейсхолдера

                bind_entry_shortcuts(entry, main_context_menu)  # Привязываем шорткаты и контекстное меню
                current_block_entries[field_key] = entry

                # Кнопка подсказки
                hint_button = ctk.CTkButton(fields_grid_frame, text="ℹ", width=28, height=28, font=("Arial", 14),
                                            command=make_show_hint_command(EXAMPLES.get(field_key, "Немає підказки"),
                                                                           field_key))
                hint_button.grid(row=i, column=2, padx=(0, 5), pady=3, sticky="e")

            # --- Автоматическое обновление зависимых полей ---
            def on_sum_or_qty_price_change(event=None, entry_map=current_block_entries):
                sum_entry = entry_map.get("сума")
                qty_entry = entry_map.get("кількість")
                price_entry = entry_map.get("ціна за одиницю")
                sum_words_entry = entry_map.get("сума прописом")
                razom_entry = entry_map.get("разом")

                # Логика расчета "Разом" из "Кількість" и "Ціна за одиницю"
                calculated_razom = None
                if qty_entry and price_entry:
                    try:
                        qty_val_str = qty_entry.get().replace(',', '.').strip()
                        price_val_str = price_entry.get().replace(',', '.').strip()
                        if qty_val_str and price_val_str:
                            qty = float(qty_val_str)
                            price = float(price_val_str)
                            calculated_razom = qty * price
                            if razom_entry and razom_entry.cget("state") == "readonly":
                                razom_entry.configure(state="normal")
                                razom_entry.set_text(f"{calculated_razom:.2f}")
                                razom_entry.configure(state="readonly")
                    except ValueError:
                        # Если ошибка конвертации, "Разом" не обновляется или очищается
                        if razom_entry and razom_entry.cget("state") == "readonly":
                            razom_entry.configure(state="normal")
                            razom_entry.set_text("")  # Очистить, если невалидные данные
                            razom_entry.configure(state="readonly")
                        calculated_razom = None  # Сброс

                # Логика обновления "Сума" и "Сума прописом"
                # Если "Разом" рассчитано, оно становится основой для "Сума"
                # Иначе, если "Сума" вводится вручную, она основа.
                source_amount_for_words = None
                if calculated_razom is not None:
                    source_amount_for_words = calculated_razom
                    # Обновляем поле "Сума", если оно есть и должно быть таким же, как "Разом"
                    if sum_entry:  # sum_entry.cget("state") != "readonly" (если оно редактируемое)
                        sum_entry.set_text(f"{calculated_razom:.2f}")
                elif sum_entry:  # Если "Разом" не рассчитано, берем значение из "Сума"
                    try:
                        sum_val_str = sum_entry.get().replace(',', '.').strip()
                        if sum_val_str:
                            source_amount_for_words = float(sum_val_str)
                    except ValueError:
                        source_amount_for_words = None

                if sum_words_entry and source_amount_for_words is not None:
                    try:
                        sum_words = number_to_ukrainian_text(source_amount_for_words)
                        sum_words_entry.configure(state="normal")
                        sum_words_entry.set_text(sum_words.capitalize())  # Первая буква заглавная
                        sum_words_entry.configure(state="readonly")
                    except Exception as e_num_to_text:
                        print(f"Error converting number to text: {e_num_to_text}")
                        sum_words_entry.configure(state="normal")
                        sum_words_entry.set_text("Помилка конвертації")
                        sum_words_entry.configure(state="readonly")
                elif sum_words_entry:  # Если нет суммы для конвертации
                    sum_words_entry.configure(state="normal")
                    sum_words_entry.set_text("")  # Очищаем
                    sum_words_entry.configure(state="readonly")

            # Привязка к полям "сума", "кількість", "ціна за одиницю"
            for key in ["сума", "кількість", "ціна за одиницю"]:
                if key in current_block_entries:
                    current_block_entries[key].bind("<KeyRelease>",
                                                    lambda event, em=current_block_entries: on_sum_or_qty_price_change(
                                                        event, em), add="+")

            # Сделать поля "сума прописом" и "разом" только для чтения изначально
            if "сума прописом" in current_block_entries:
                current_block_entries["сума прописом"].configure(state="readonly")
            if "разом" in current_block_entries:
                current_block_entries["разом"].configure(state="readonly")

            # Вызываем один раз для инициализации, если есть значения
            on_sum_or_qty_price_change(entry_map=current_block_entries)

            # --- Кнопки управления блоком ---
            block_actions_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
            block_actions_frame.pack(fill="x", padx=5, pady=5)

            # Сохраняем ссылку на фрейм и путь в словаре блока
            # Этот block_dict будет добавлен в document_blocks
            block_dict_for_list = {"path": template_filepath, "entries": current_block_entries, "frame": block_frame}

            def clear_block_fields():
                if messagebox.askokcancel("Очистити поля", "Ви впевнені, що хочете очистити всі поля цього договору?"):
                    for entry_widget in current_block_entries.values():
                        if hasattr(entry_widget, 'clear'):  # Используем наш метод clear у CustomEntry
                            entry_widget.clear()
                        else:  # Стандартная очистка для других виджетов, если вдруг появятся
                            entry_widget.configure(state="normal")
                            entry_widget.delete(0, 'end')
                    # После очистки, пересчитываем зависимые поля
                    on_sum_or_qty_price_change(entry_map=current_block_entries)

            def replace_block_template():
                nonlocal block_dict_for_list  # Для изменения path в словаре
                new_filepath = filedialog.askopenfilename(
                    title="Оберіть новий шаблон для цього блоку",
                    filetypes=[("Word Documents", "*.docm *.docx")]
                )
                if new_filepath and new_filepath != block_dict_for_list["path"]:
                    # Проверяем, не дублируется ли новый путь с другими блоками
                    for existing_block in document_blocks:
                        if existing_block["path"] == new_filepath and existing_block is not block_dict_for_list:
                            messagebox.showwarning("Увага",
                                                   f"Шаблон '{os.path.basename(new_filepath)}' вже додано в іншому блоці.")
                            return

                    block_dict_for_list["path"] = new_filepath
                    path_label.configure(text=f"Шаблон: {os.path.basename(new_filepath)} ({new_filepath})")
                    # Можно добавить логику перезагрузки памяти для нового шаблона, если это нужно
                    new_template_memory = get_template_memory(new_filepath)
                    general_mem = load_memory()
                    for f_key, entry_w in current_block_entries.items():
                        new_val = new_template_memory.get(f_key, general_mem.get(f_key))
                        entry_w.set_text(new_val if new_val is not None else "")
                    on_sum_or_qty_price_change(entry_map=current_block_entries)  # Обновить зависимые поля
                    messagebox.showinfo("Шаблон замінено",
                                        f"Шаблон для цього блоку замінено на\n'{os.path.basename(new_filepath)}'.\nДані спробували завантажити з пам'яті.")

            def remove_this_block():
                nonlocal block_dict_for_list, block_frame  # Для переменных из create_document_fields_block
                global document_blocks  # Для глобальной переменной

                if messagebox.askokcancel("Видалити договір", "Ви впевнені, що хочете видалити цей блок договору?"):
                    if block_dict_for_list in document_blocks:  # Читаем block_dict_for_list, читаем document_blocks
                        document_blocks.remove(block_dict_for_list)  # Изменяем document_blocks
                    block_frame.destroy()  # Используем block_frame
                    # Обновить скроллбар, если необходимо (обычно CTkScrollableFrame делает это автоматически)

            ctk.CTkButton(block_actions_frame, text="🧹 Очистити поля", command=clear_block_fields).pack(side="left",
                                                                                                        padx=3)
            ctk.CTkButton(block_actions_frame, text="🔄 Замінити шаблон", command=replace_block_template).pack(
                side="left", padx=3)

            # Кнопка удаления справа
            remove_button = ctk.CTkButton(header_frame, text="🗑", width=28, height=28, fg_color="gray50",
                                          hover_color="gray40", command=remove_this_block)
            remove_button.pack(side="right", padx=(5, 0))

            document_blocks.append(block_dict_for_list)
            # Прокрутка к новому блоку, если он не виден (необязательно, но удобно)
            # scroll_frame_main._parent_canvas.yview_moveto(1) # Прокрутка вниз

        # Загрузка сохраненных блоков при запуске (если это нужно)
        # Это более сложная логика, так как нужно будет восстанавливать GUI для каждого сохраненного пути
        # Для простоты, пока оставим так, что пользователь добавляет шаблоны каждый раз
        # или можно загрузить последний использованный набор шаблонов.

        # --- Глобальные горячие клавиши ---
        # main_app_root.bind_all("<Control-s>", lambda event: export_document_data_to_excel(document_blocks, FIELDS)) # Уже есть кнопка
        main_app_root.bind_all("<Control-g>", lambda event: combined_generation_process())
        main_app_root.bind_all("<Control-n>", lambda event: add_new_template_block())  # Ctrl+N для нового договора
        main_app_root.bind_all("<F1>", lambda event: messagebox.showinfo("Допомога",
                                                                         "SportForAll Document Generator\n\n- Додайте шаблони договорів.\n- Заповніть поля.\n- Натисніть 'Згенерувати ВСІ документи'."))

        main_app_root.mainloop()

    except Exception as e_launch:  # Ловим ошибки на уровне запуска главного приложения
        log_and_show_error(type(e_launch), e_launch, sys.exc_info()[2])
        # messagebox.showerror("Критична помилка запуску", f"Не вдалося запустити основний додаток:\n{e_launch}")
        # input("Натисніть Enter для завершення...") # log_and_show_error это сделает


# --- Точка входа в программу ---
if __name__ == "__main__":
    # Сначала запрашиваем пароль, потом запускаем основное приложение
    # ask_password передает launch_main_app как колбэк, который будет вызван при успешном вводе пароля.
    try:
        ask_password(on_success_callback=launch_main_app)
    except Exception as e_main_start:
        # Эта ошибка будет поймана, если сам ask_password или launch_main_app выбросят исключение
        # которое не было обработано внутри них или глобальным обработчиком (что маловероятно для launch_main_app)
        # Глобальный обработчик должен поймать большинство ошибок из launch_main_app.
        # Ошибки из ask_password могут быть пойманы здесь, если они произойдут до его mainloop.
        log_and_show_error(type(e_main_start), f"Критична помилка при старті програми: {e_main_start}",
                           sys.exc_info()[2])
        # input("Натисніть Enter для завершення...") # log_and_show_error это сделает
        sys.exit(1)