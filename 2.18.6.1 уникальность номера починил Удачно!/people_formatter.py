# people_formatter.py

from globals import SPECIAL_ROLES, PEOPLE_CONFIG


def generate_replacements(selected_people,
                          get_special_role_func,
                          generate_people_list_text_func,
                          detect_used_part_placeholders_func,
                          invisible_char,
                          text=None):
    """Генерує словник замін для шаблонів"""
    print(f"[DEBUG] Generating replacements for selected people: {selected_people}")
    replacements = {}

    # 1. Плейсхолдери для всіх людей – пусті
    for person_id, person_data in PEOPLE_CONFIG.items():
        replacements[person_data['placeholders']['full_block']] = ""
        replacements[person_data['placeholders']['name_only']] = ""
        replacements[person_data['placeholders']['position_only']] = ""
        # Додаємо accusative блоки - також спочатку пусті
        if 'accusative_block' in person_data['placeholders']:
            replacements[person_data['placeholders']['accusative_block']] = ""

    # 2. Плейсхолдери частин списку – невидимий символ
    if text:
        used_part_numbers = detect_used_part_placeholders_func(text)
    else:
        used_part_numbers = set(range(1, 11))

    for i in used_part_numbers:
        replacements[f"{{{{SELECTED_PEOPLE_PART_{i}}}}}"] = invisible_char

    replacements["{{SELECTED_PEOPLE_LIST}}"] = invisible_char
    replacements["{{SELECTED_PEOPLE_PARTS_COUNT}}"] = "0"

    for role_config in SPECIAL_ROLES.values():
        replacements[role_config['placeholder']] = invisible_char

    # 5. Список людей
    people_list_text = generate_people_list_text_func()
    if people_list_text:
        if len(people_list_text) > 150:
            print(f"[DEBUG] Long text ({len(people_list_text)} chars), splitting into parts")
            parts = people_list_text.split(', ')
            chunks = []
            current_chunk = ""
            for part in parts:
                test_chunk = current_chunk + (", " if current_chunk else "") + part
                if len(test_chunk) <= 150:
                    current_chunk = test_chunk
                else:
                    if current_chunk:
                        chunks.append(current_chunk)
                    current_chunk = part
            if current_chunk:
                chunks.append(current_chunk)

            print(f"[DEBUG] Split into {len(chunks)} chunks")
            filled_parts = 0
            for i, chunk in enumerate(chunks):
                part_number = i + 1
                placeholder = f"{{{{SELECTED_PEOPLE_PART_{part_number}}}}}"
                if part_number in used_part_numbers:
                    replacements[placeholder] = chunk if i == 0 else f", {chunk}"
                    filled_parts += 1
                    print(f"[DEBUG] Set {placeholder}: '{replacements[placeholder][:50]}...'")
            replacements["{{SELECTED_PEOPLE_PARTS_COUNT}}"] = str(filled_parts)
        else:
            if 1 in used_part_numbers:
                replacements["{{SELECTED_PEOPLE_PART_1}}"] = people_list_text
                replacements["{{SELECTED_PEOPLE_PARTS_COUNT}}"] = "1"
                print(f"[DEBUG] Set short text to PART_1: '{people_list_text}'")
            else:
                print("[DEBUG] PART_1 not used in text, keeping invisible char")
    else:
        print("[DEBUG] No people selected, all PART placeholders get invisible chars")

    # 6. Індивідуальні плейсхолдери обраних людей
    for person_id in selected_people:
        person_data = PEOPLE_CONFIG.get(person_id)
        if person_data:
            full_block = (
                f"{person_data['position']}\r\n\r\n"
                f"1___1________________2025 року \t\t\t\t{person_data['name']}\r\n"
            )
            replacements[person_data['placeholders']['full_block']] = full_block
            replacements[person_data['placeholders']['name_only']] = person_data['name']
            replacements[person_data['placeholders']['position_only']] = person_data['position']

            # Додаємо обробку accusative блоків
            if 'accusative_block' in person_data['placeholders'] and 'accusative_block' in person_data:
                replacements[person_data['placeholders']['accusative_block']] = person_data['accusative_block']
                print(f"[DEBUG] Set accusative block for {person_id}: '{person_data['accusative_block']}'")

            print(f"[DEBUG] Set placeholders for selected person: {person_id}")

    # 7. Спеціальні ролі
    for role_id, role_config in SPECIAL_ROLES.items():
        selected_person_id = get_special_role_func(role_id)
        person_data = PEOPLE_CONFIG.get(selected_person_id)
        if person_data:
            if role_id == "material_responsible":
                material_block = (
                    f"Матеріально-відповідальна особа:\r\n\r\n"
                    f"{person_data['position']}\r\n\r\n"
                    f"1___1________________2025 року \t\t\t\t{person_data['name']}\r\n"
                )
                replacements[role_config['placeholder']] = material_block
            else:
                replacements[role_config['placeholder']] = person_data['name']
            print(f"[DEBUG] Set special role placeholder: {role_id}")

    return replacements
