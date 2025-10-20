# state_manager.py

from globals import document_blocks
from data_persistence import save_full_state, load_saved_state



def save_current_state(blocks=None, tabview=None):
    """Зберігає поточний стан програми"""
    if blocks is None:
        blocks = document_blocks

    def safe_entry_value(entry):
        try:
            if hasattr(entry, "get") and callable(entry.get):
                return entry.get()
            elif callable(entry):
                return str(entry())
            else:
                return str(entry)
        except Exception:
            return f"[Unserializable:{type(entry).__name__}]"

    simplified_blocks = []
    for block in blocks:
        simplified_block = {
            "path": block["path"],
            "tab_name": block.get("tab_name", ""),
            "entries": {
                key: safe_entry_value(entry) for key, entry in block["entries"].items()
            }
        }
        simplified_blocks.append(simplified_block)

    tabs = list(tabview._tab_dict.keys()) if tabview else []

    app_state = {
        "tabs": tabs,
        "document_blocks": simplified_blocks
    }

    save_full_state(app_state)


def restore_saved_state(tabview, create_document_fields_block):
    """Відновлює вкладки та блоки договорів"""
    from events import add_event  # 🔁 Импорт внутри функции

    """Відновлює вкладки та блоки договорів"""
    saved_state = load_saved_state()
    if not saved_state:
        return

    for tab_name in saved_state.get("tabs", []):
        if tab_name not in tabview._tab_dict:
            add_event(tab_name, tabview, restore=True)

    for block_data in saved_state.get("document_blocks", []):
        tab_name = block_data.get("tab_name")
        if not tab_name or tab_name not in tabview._tab_dict:
            continue
        try:
            tab = tabview.tab(tab_name)
            tabview.set(tab_name)
            create_document_fields_block(tab.contracts_frame, tabview, template_filepath=block_data.get("path"))
            if document_blocks:
                last_block = document_blocks[-1]
                for key, value in block_data.get("entries", {}).items():
                    entry = last_block["entries"].get(key)
                    if entry:
                        entry.set_text(value)
        except Exception as e:
            print(f"Помилка при завантаженні блоку для вкладки '{tab_name}': {e}")

