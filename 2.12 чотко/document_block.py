# document_block.py

import os
import tkinter.messagebox as messagebox
import customtkinter as ctk
from tkinter import filedialog

from globals import FIELDS, EXAMPLES, document_blocks
from gui_utils import bind_entry_shortcuts, create_context_menu
from custom_widgets import CustomEntry
from text_utils import number_to_ukrainian_text
from data_persistence import get_template_memory, load_memory
from state_manager import save_current_state


def create_document_fields_block(parent_frame, tabview, template_filepath=None):
    if not template_filepath:
        template_filepath = filedialog.askopenfilename(
            title="Оберіть шаблон договору",
            filetypes=[("Word Documents", "*.docm *.docx")]
        )
        if not template_filepath:
            return

    block_frame = ctk.CTkFrame(parent_frame, border_width=1, border_color="gray70")
    block_frame.pack(pady=10, padx=5, fill="x")

    header_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    header_frame.pack(fill="x", padx=5, pady=(5, 0))

    path_label = ctk.CTkLabel(header_frame,
                              text=f"Шаблон: {os.path.basename(template_filepath)} ({template_filepath})",
                              text_color="blue", anchor="w", wraplength=800)
    path_label.pack(side="left", padx=(0, 5), expand=True, fill="x")

    current_block_entries = {}
    template_specific_memory = get_template_memory(template_filepath)
    general_memory = load_memory()

    # ✅ Создаём отдельный фрейм для .grid() элементов
    fields_grid_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    fields_grid_frame.pack(fill="both", expand=True, padx=5, pady=5)

    main_context_menu = create_context_menu(block_frame)

    for i, field_key in enumerate(FIELDS):
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

        if field_key == "захід":
            entry.set_text(tabview.get())

            def on_event_name_change(event):
                new_name = entry.get().strip()
                if new_name:
                    old_name = tabview.get()
                    try:
                        tabview._segmented_button._update_button(new_name, old_name)
                        tabview._name_to_tab_dict[new_name] = tabview._name_to_tab_dict.pop(old_name)
                    except Exception as e:
                        print("Помилка перейменування вкладки:", e)

            entry.bind("<FocusOut>", on_event_name_change)

        hint_button = ctk.CTkButton(fields_grid_frame, text="ℹ", width=28, height=28, font=("Arial", 14),
                                    command=lambda h=EXAMPLES.get(field_key, "Немає підказки"), f=field_key:
                                    messagebox.showinfo(f"Підказка для <{f}>", h))
        hint_button.grid(row=i, column=2, padx=(0, 5), pady=3, sticky="e")

    def on_sum_or_qty_price_change(event=None):
        qty_entry = current_block_entries.get("кількість")
        price_entry = current_block_entries.get("ціна за одиницю")
        sum_entry = current_block_entries.get("сума")
        sum_words_entry = current_block_entries.get("сума прописом")
        razom_entry = current_block_entries.get("разом")
        zagalna_entry = current_block_entries.get("загальна сума")

        try:
            qty = float(qty_entry.get().replace(",", "."))
            price = float(price_entry.get().replace(",", "."))
            total = qty * price
            if razom_entry:
                razom_entry.configure(state="normal")
                razom_entry.set_text(f"{total:.2f}")
                razom_entry.configure(state="readonly")
            if zagalna_entry:
                zagalna_entry.configure(state="normal")
                zagalna_entry.set_text(f"{total:.2f}")
                zagalna_entry.configure(state="readonly")
            if sum_entry:
                sum_entry.set_text(f"{int(total)} грн {int((total - int(total)) * 100):02d} коп.")
            if sum_words_entry:
                words = number_to_ukrainian_text(total).capitalize()
                sum_words_entry.configure(state="normal")
                sum_words_entry.set_text(words)
                sum_words_entry.configure(state="readonly")
        except:
            pass

    for key in ["сума", "кількість", "ціна за одиницю"]:
        if key in current_block_entries:
            current_block_entries[key].bind("<KeyRelease>", on_sum_or_qty_price_change, add="+")

    for key in ["сума прописом", "разом", "загальна сума"]:
        if key in current_block_entries:
            current_block_entries[key].configure(state="readonly")

    on_sum_or_qty_price_change()

    block_actions_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    block_actions_frame.pack(fill="x", padx=5, pady=5)

    block_dict = {
        "path": template_filepath,
        "entries": current_block_entries,
        "frame": block_frame,
        "tab_name": tabview.get()
    }

    def clear_block_fields():
        if messagebox.askokcancel("Очистити поля", "Очистити всі поля цього договору?"):
            for entry in current_block_entries.values():
                entry.configure(state="normal")
                entry.set_text("")
            on_sum_or_qty_price_change()

    def replace_block_template():
        new_path = filedialog.askopenfilename(
            title="Оберіть новий шаблон",
            filetypes=[("Word Documents", "*.docm *.docx")]
        )
        if new_path:
            block_dict["path"] = new_path
            path_label.configure(text=f"Шаблон: {os.path.basename(new_path)} ({new_path})")
            new_mem = get_template_memory(new_path)
            general = load_memory()
            for key, entry in current_block_entries.items():
                val = new_mem.get(key, general.get(key))
                entry.set_text(val if val else "")
            on_sum_or_qty_price_change()
            messagebox.showinfo("Шаблон замінено", f"Шаблон замінено на {os.path.basename(new_path)}")

    def remove_this_block():
        if messagebox.askokcancel("Видалити", "Видалити цей блок договору?"):
            if block_dict in document_blocks:
                document_blocks.remove(block_dict)
            block_frame.destroy()

    ctk.CTkButton(block_actions_frame, text="🧹 Очистити поля", command=clear_block_fields).pack(side="left", padx=3)
    ctk.CTkButton(block_actions_frame, text="🔄 Замінити шаблон", command=replace_block_template).pack(side="left", padx=3)

    remove_button = ctk.CTkButton(header_frame, text="🗑", width=28, height=28, fg_color="gray50", hover_color="gray40",
                                  command=remove_this_block)
    remove_button.pack(side="right", padx=(5, 0))

    document_blocks.append(block_dict)
    save_current_state(document_blocks, tabview)
