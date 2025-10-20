# sportforall/utils.py

def number_to_currency_text(number: float | int | str | None) -> str:
    """
    Преобразует числовое значение (или строку, или None) в текстовое представление суммы (гривны).
    Пример: 15000.50 -> "П'ятнадцять тисяч гривень 50 копійок"
    Возвращает "Некоректна сума", если ввод не является числом.
    """
    # TODO: Implement actual number to text conversion logic for Ukrainian
    # This requires a more complex implementation or a library.
    # For now, return a simplified placeholder string.
    if number is None or str(number).strip() == "":
        number = 0.0
    try:
        # Попытаемся преобразовать ввод в число, обрабатывая строки с запятой
        numeric_value = float(str(number).replace(",", "."))
        numeric_value = round(numeric_value, 2) # Округляем до двух знаков после запятой

        integer_part = int(numeric_value)
        fractional_part = round((numeric_value - integer_part) * 100)

        # TODO: Реализовать склонение гривен и копеек, перевод чисел в текст
        # Пока заглушка:
        grn_text = "гривень" # Нужна правильная форма
        kop_text = "копійок" # Нужна правильная форма

        # Очень упрощенная текстовая часть (только число цифрами)
        integer_text = str(integer_part) # Заменить на перевод числа в текст
        fractional_text = f"{fractional_part:02d}" # Форматируем копейки двумя цифрами

        return f"{integer_text} {grn_text} {fractional_text} {kop_text}"

    except (ValueError, TypeError):
        return "Некоректна сума"

# Вы можете добавить другие вспомогательные функции сюда в будущем