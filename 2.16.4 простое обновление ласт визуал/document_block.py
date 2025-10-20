# document_block.py 

import os
import tkinter.messagebox as messagebox
import customtkinter as ctk
from tkinter import filedialog

from globals import EXAMPLES, document_blocks
from gui_utils import bind_entry_shortcuts, create_context_menu
from custom_widgets import CustomEntry
from text_utils import number_to_ukrainian_text
from data_persistence import get_template_memory, load_memory
from state_manager import save_current_state
from event_common_fields import fill_common_fields_for_new_contract, COMMON_FIELDS
from generation import extract_placeholders_from_word

# Глобальные поля мероприятия (создаются один раз для всего мероприятия)
global_event_fields = {}
global_event_frame = None


def create_global_event_fields(parent_frame, tabview):
    """Создает глобальные поля мероприятия один раз для всей вкладки"""
    global global_event_fields, global_event_frame

    # Если поля уже созданы для этой вкладки, не создаем повторно
    current_tab = tabview.get()
    if global_event_frame and current_tab in global_event_fields:
        return global_event_fields[current_tab]

    # Создаем фрейм для глобальных полей мероприятия
    event_frame = ctk.CTkFrame(parent_frame, border_width=2, border_color="#2E8B57")
    event_frame.pack(pady=(5, 10), padx=5, fill="x")

    # Заголовок
    header_label = ctk.CTkLabel(event_frame, text="🎯 ГЛОБАЛЬНЫЕ ДАННЫЕ МЕРОПРИЯТИЯ",
                                font=("Arial", 16, "bold"), text_color="#2E8B57")
    header_label.pack(pady=(10, 5))

    # Сетка для полей
    grid_frame = ctk.CTkFrame(event_frame, fg_color="transparent")
    grid_frame.pack(fill="x", padx=15, pady=(0, 15))

    # Поля мероприятия
    event_fields_data = [
        ("захід", "Название мероприятия"),
        ("дата", "Дата проведения мероприятия"),
        ("адреса", "Адрес проведения мероприятия")
    ]

    current_event_entries = {}
    general_memory = load_memory()
    main_context_menu = create_context_menu(event_frame)

    for i, (field_key, description) in enumerate(event_fields_data):
        # Лейбл
        label = ctk.CTkLabel(grid_frame, text=f"<{field_key}>", anchor="w", width=100,
                             font=("Arial", 13, "bold"), text_color="#2E8B57")
        label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

        # Поле ввода
        entry = CustomEntry(grid_frame, field_name=field_key, examples_dict=EXAMPLES)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        grid_frame.columnconfigure(1, weight=1)

        # Заполняем сохраненными данными
        saved_value = general_memory.get(field_key)
        if saved_value:
            entry.set_text(saved_value)

        bind_entry_shortcuts(entry, main_context_menu)
        current_event_entries[field_key] = entry

        # Кнопка подсказки
        hint_text = EXAMPLES.get(field_key, description)
        hint_button = ctk.CTkButton(grid_frame, text="ℹ", width=30, height=30, font=("Arial", 14),
                                    command=lambda h=hint_text, f=field_key:
                                    messagebox.showinfo(f"Подсказка для <{f}>", h))
        hint_button.grid(row=i, column=2, padx=(5, 0), pady=5, sticky="e")

    # Сохраняем ссылки на глобальные поля
    if current_tab not in global_event_fields:
        global_event_fields[current_tab] = {}

    global_event_fields[current_tab] = current_event_entries
    global_event_frame = event_frame

    # Заполняем общие поля из данных мероприятия
    fill_common_fields_for_new_contract(current_tab, current_event_entries)

    return current_event_entries


def get_global_event_fields(tab_name):
    """Возвращает глобальные поля мероприятия для указанной вкладки"""
    return global_event_fields.get(tab_name, {})


def create_document_fields_block(parent_frame, tabview=None, template_path=None):
    if not tabview:
        messagebox.showerror("Помилка", "TabView не переданий")
        return

    if not template_path:
        template_path = filedialog.askopenfilename(
            title="Оберіть шаблон договору",
            filetypes=[("Word Documents", "*.docm *.docx")]
        )
        if not template_path:
            return

    try:
        dynamic_fields = extract_placeholders_from_word(template_path)
        print(f"[DEBUG] Знайдені поля в шаблоні {os.path.basename(template_path)}: {dynamic_fields}")

        if not dynamic_fields:
            messagebox.showwarning("Увага",
                                   f"У шаблоні {os.path.basename(template_path)} не знайдено жодного плейсхолдера типу <поле>.\n"
                                   "Переконайтеся, що у шаблоні є поля у форматі <назва_поля>")
            return

    except Exception as e:
        messagebox.showerror("Помилка", f"Не вдалося прочитати шаблон:\n{e}")
        return

    # Создаем или получаем глобальные поля мероприятия
    global_entries = create_global_event_fields(parent_frame, tabview)

    # Создаем блок для полей конкретного договора
    block_frame = ctk.CTkFrame(parent_frame, border_width=1, border_color="gray70")
    block_frame.pack(pady=10, padx=5, fill="x")

    header_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    header_frame.pack(fill="x", padx=5, pady=(5, 0))

    path_label = ctk.CTkLabel(header_frame,
                              text=f"Шаблон: {os.path.basename(template_path)} ({len(dynamic_fields)} полів)",
                              text_color="blue", anchor="w", wraplength=800)
    path_label.pack(side="left", padx=(0, 5), expand=True, fill="x")

    current_block_entries = {}
    template_specific_memory = get_template_memory(template_path)
    general_memory = load_memory()

    # Определяем поля договора (исключаем глобальные поля мероприятия)
    global_event_field_names = ["захід", "дата", "адреса"]
    contract_fields = [field for field in dynamic_fields if field not in global_event_field_names]

    # Добавляем остальные общие поля (если есть в шаблоне)
    template_common_fields = [field for field in contract_fields if field in COMMON_FIELDS]

    # Создаем контекстное меню
    main_context_menu = create_context_menu(block_frame)

    # БЛОК ОБЩИХ ПОЛЕЙ ДОГОВОРА (кроме глобальных полей мероприятия)
    if template_common_fields:
        common_data_frame = ctk.CTkFrame(block_frame)
        common_data_frame.pack(fill="x", padx=5, pady=5)

        common_label = ctk.CTkLabel(common_data_frame, text="📋 Общие данные договора",
                                    font=("Arial", 14, "bold"), text_color="#FF6B35")
        common_label.pack(pady=(10, 5))

        common_grid_frame = ctk.CTkFrame(common_data_frame, fg_color="transparent")
        common_grid_frame.pack(fill="x", padx=10, pady=(0, 10))

        for i, field_key in enumerate(template_common_fields):
            label = ctk.CTkLabel(common_grid_frame, text=f"<{field_key}>", anchor="w", width=140,
                                 font=("Arial", 12, "bold"), text_color="#FF6B35")
            label.grid(row=i, column=0, padx=5, pady=3, sticky="w")

            entry = CustomEntry(common_grid_frame, field_name=field_key, examples_dict=EXAMPLES)
            entry.grid(row=i, column=1, padx=5, pady=3, sticky="ew")
            common_grid_frame.columnconfigure(1, weight=1)

            saved_value = template_specific_memory.get(field_key, general_memory.get(field_key))
            if saved_value is not None:
                entry.set_text(saved_value)

            bind_entry_shortcuts(entry, main_context_menu)
            current_block_entries[field_key] = entry

            hint_text = EXAMPLES.get(field_key, f"Общее поле договора: {field_key}")
            hint_button = ctk.CTkButton(common_grid_frame, text="ℹ", width=28, height=28, font=("Arial", 14),
                                        command=lambda h=hint_text, f=field_key:
                                        messagebox.showinfo(f"Підказка для <{f}>", h))
            hint_button.grid(row=i, column=2, padx=(0, 5), pady=3, sticky="e")

    # БЛОК СПЕЦИФИЧЕСКИХ ПОЛЕЙ ДОГОВОРА
    fields_grid_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    fields_grid_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # Фильтруем поля - показываем только специфические поля договора
    specific_contract_fields = [field for field in contract_fields if field not in COMMON_FIELDS]

    # Сортируем поля по приоритету
    priority_fields = ["товар", "назва", "найменування", "предмет", "послуга", "кількість", "ціна за одиницю", "сума",
                       "разом"]
    sorted_fields = []

    for priority_field in priority_fields:
        if priority_field in specific_contract_fields:
            sorted_fields.append(priority_field)
            specific_contract_fields.remove(priority_field)

    sorted_fields.extend(sorted(specific_contract_fields))

    for i, field_key in enumerate(sorted_fields):
        label = ctk.CTkLabel(fields_grid_frame, text=f"<{field_key}>", anchor="w", width=140, font=("Arial", 12))
        label.grid(row=i, column=0, padx=5, pady=3, sticky="w")

        entry = CustomEntry(fields_grid_frame, field_name=field_key, examples_dict=EXAMPLES)
        entry.grid(row=i, column=1, padx=5, pady=3, sticky="ew")
        fields_grid_frame.columnconfigure(1, weight=1)

        saved_value = template_specific_memory.get(field_key, general_memory.get(field_key))
        if saved_value is not None:
            entry.set_text(saved_value)

        bind_entry_shortcuts(entry, main_context_menu)
        current_block_entries[field_key] = entry

        hint_text = EXAMPLES.get(field_key, f"Поле для заповнення: {field_key}")
        hint_button = ctk.CTkButton(fields_grid_frame, text="ℹ", width=28, height=28, font=("Arial", 14),
                                    command=lambda h=hint_text, f=field_key:
                                    messagebox.showinfo(f"Підказка для <{f}>", h))
        hint_button.grid(row=i, column=2, padx=(0, 5), pady=3, sticky="e")

    # АВТОМАТИЧЕСКИЕ РАСЧЕТЫ
    def on_sum_or_qty_price_change(event=None):
        try:
            # Безпечне отримання значення кількості
            qty = 0
            if "кількість" in current_block_entries:
                qty_text = current_block_entries["кількість"].get().replace(",", ".")
                qty = float(qty_text) if qty_text.strip() else 0

            # Безпечне отримання значення ціни
            price = 0
            if "ціна за одиницю" in current_block_entries:
                price_text = current_block_entries["ціна за одиницю"].get().replace(",", ".")
                price = float(price_text) if price_text.strip() else 0

            total = qty * price

            # Оновлення поля "разом"
            if "разом" in current_block_entries:
                entry = current_block_entries["разом"]
                entry.configure(state="normal")
                entry.set_text(f"{total:.2f}")
                entry.configure(state="readonly")

            # Оновлення поля "загальна сума"
            if "загальна сума" in current_block_entries:
                entry = current_block_entries["загальна сума"]
                entry.configure(state="normal")
                entry.set_text(f"{total:.2f}")
                entry.configure(state="readonly")

            # Оновлення поля "сума" (використовуємо старий код)
            sum_entry = current_block_entries.get("сума")
            if sum_entry:
                sum_entry.set_text(f"{int(total)} грн {int((total - int(total)) * 100):02d} коп.")

            # Оновлення поля "сума прописом" (не змінюємо - воно правильне)
            if "сума прописом" in current_block_entries:
                entry = current_block_entries["сума прописом"]
                entry.configure(state="normal")
                entry.set_text(number_to_ukrainian_text(total).capitalize())
                entry.configure(state="readonly")

        except ValueError as e:
            print(f"[DEBUG] Помилка конвертації числа: {e}")
        except Exception as e:
            print(f"[DEBUG] Помилка в розрахунках: {e}")

    # Функція для автоформатування ціни
    def format_price_field(event=None):
        """Автоматично форматує поле 'ціна за одиницю' до формату 0.00"""
        if "ціна за одиницю" not in current_block_entries:
            return

        try:
            price_entry = current_block_entries["ціна за одиницю"]
            price_text = price_entry.get().replace(",", ".").strip()

            if price_text and price_text != "":
                # Перевіряємо чи це валідне число
                price_value = float(price_text)
                # Форматуємо до двох десяткових знаків
                formatted_price = f"{price_value:.2f}"
                price_entry.set_text(formatted_price)
        except ValueError:
            # Якщо не вдається конвертувати - залишаємо як є
            pass

    # ПРИВ'ЯЗКИ ПОДІЙ для автоматичних розрахунків
    if "кількість" in current_block_entries:
        current_block_entries["кількість"].bind("<KeyRelease>", on_sum_or_qty_price_change, add="+")
        current_block_entries["кількість"].bind("<FocusOut>", on_sum_or_qty_price_change, add="+")

    if "ціна за одиницю" in current_block_entries:
        # Для розрахунків
        current_block_entries["ціна за одиницю"].bind("<KeyRelease>", on_sum_or_qty_price_change, add="+")
        # Для автоформатування ціни (спрацьовує коли користувач покидає поле)
        current_block_entries["ціна за одиницю"].bind("<FocusOut>",
                                                      lambda e: [format_price_field(e), on_sum_or_qty_price_change(e)],
                                                      add="+")

    # Робимо readonly поля, які автоматично обчислюються
    readonly_fields = ["сума прописом", "разом", "загальна сума"]
    for key in readonly_fields:
        if key in current_block_entries:
            current_block_entries[key].configure(state="readonly")

    # Викликаємо початковий розрахунок
    on_sum_or_qty_price_change()

    # ПАНЕЛЬ ДЕЙСТВИЙ ДЛЯ БЛОКА
    block_actions_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    block_actions_frame.pack(fill="x", padx=5, pady=5)

    def clear_block_fields():
        if messagebox.askokcancel("Очистити поля", "Очистити всі поля цього договору?"):
            for field_key, entry in current_block_entries.items():
                # Не очищаем общие поля
                if field_key not in COMMON_FIELDS:
                    entry.configure(state="normal")
                    entry.set_text("")
            on_sum_or_qty_price_change()

    def replace_block_template():
        new_path = filedialog.askopenfilename(
            title="Оберіть новий шаблон",
            filetypes=[("Word Documents", "*.docm *.docx")]
        )
        if new_path:
            try:
                new_placeholders = extract_placeholders_from_word(new_path)
                if not new_placeholders:
                    messagebox.showwarning("Увага", "У новому шаблоні не знайдено плейсхолдерів!")
                    return

                path_label.configure(text=f"Шаблон: {os.path.basename(new_path)} ({len(new_placeholders)} полів)")
                messagebox.showinfo("Увага", "Щоб застосувати новий шаблон, видаліть блок і створіть новий.")
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося прочитати новий шаблон:\n{e}")

    def remove_this_block():
        if messagebox.askokcancel("Видалити", "Видалити цей блок договору?"):
            if block_dict in document_blocks:
                document_blocks.remove(block_dict)
            block_frame.destroy()
            save_current_state(document_blocks, tabview)

    # Кнопки действий
    ctk.CTkButton(block_actions_frame, text="🧹 Очистити поля", command=clear_block_fields).pack(side="left", padx=3)
    ctk.CTkButton(block_actions_frame, text="🔄 Замінити шаблон", command=replace_block_template).pack(side="left",
                                                                                                      padx=3)

    # Информационная метка о полях
    fields_list = sorted(list(dynamic_fields))
    info_text = f"Знайдено {len(dynamic_fields)} полів: " + ", ".join(fields_list[:3])
    if len(dynamic_fields) > 3:
        info_text += f" та ще {len(dynamic_fields) - 3}..."

    info_label = ctk.CTkLabel(block_actions_frame, text=info_text, text_color="gray60", font=("Arial", 10))
    info_label.pack(side="left", padx=10)

    # Кнопка удаления в хедере
    remove_button = ctk.CTkButton(header_frame, text="🗑", width=28, height=28, fg_color="gray50", hover_color="gray40",
                                  command=remove_this_block)
    remove_button.pack(side="right", padx=(5, 0))

    # СОХРАНЕНИЕ БЛОКА
    # Объединяем глобальные поля мероприятия с полями договора для сохранения
    all_entries = {}
    all_entries.update(global_entries)  # Добавляем глобальные поля
    all_entries.update(current_block_entries)  # Добавляем поля договора

    block_dict = {
        "path": template_path,
        "entries": all_entries,  # Все поля (глобальные + договора)
        "contract_entries": current_block_entries,  # Только поля договора
        "frame": block_frame,
        "tab_name": tabview.get(),
        "placeholders": dynamic_fields
    }

    document_blocks.append(block_dict)
    save_current_state(document_blocks, tabview)

    print(
        f"[INFO] Створено блок договору з {len(dynamic_fields)} полями для шаблону: {os.path.basename(template_path)}")
    print(f"[INFO] Глобальні поля мероприятия: захід, дата, адреса")
    if template_common_fields:
        print(f"[INFO] Загальні поля договора: {template_common_fields}")