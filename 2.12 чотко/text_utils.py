# text_utils.py

def number_to_ukrainian_text(amount):
    """Функция конвертации числа в текст украинским языком"""
    UNITS = (
        ("нуль",),
        ("одна", "один"),
        ("дві", "два"),
        ("три", "три"),
        ("чотири", "чотири"),
        ("п’ять",),
        ("шість",),
        ("сім",),
        ("вісім",),
        ("дев’ять",),
    )

    TENS = (
        "", "десять", "двадцять", "тридцять", "сорок", "п’ятдесят",
        "шістдесят", "сімдесят", "вісімдесят", "дев’яносто"
    )

    TEENS = (
        "десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять",
        "п’ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев’ятнадцять"
    )

    HUNDREDS = (
        "", "сто", "двісті", "триста", "чотириста",
        "п’ятсот", "шістсот", "сімсот", "вісімсот", "дев’ятсот"
    )

    ORDERS = [
        (("гривня", "гривні", "гривень"), "f"),
        (("тисяча", "тисячі", "тисяч"), "f"),
        (("мільйон", "мільйони", "мільйонів"), "m"),
        (("мільярд", "мільярди", "мільярдів"), "m")
    ]

    def get_plural_form(n):
        if n % 10 == 1 and n % 100 != 11:
            return 0
        elif 2 <= n % 10 <= 4 and not (12 <= n % 100 <= 14):
            return 1
        return 2

    def convert_triplet(n, gender):
        result = []
        n = int(n)
        if n == 0:
            return []

        hundreds = n // 100
        tens_units = n % 100
        tens = tens_units // 10
        units = tens_units % 10

        result.append(HUNDREDS[hundreds])

        if tens_units >= 10 and tens_units < 20:
            result.append(TEENS[tens_units - 10])
        else:
            result.append(TENS[tens])
            if units > 0:
                if units < 3:
                    result.append(UNITS[units][0 if gender == "f" else 1])
                else:
                    result.append(UNITS[units][0])

        return [word for word in result if word]

    integer_part = int(amount)
    fractional_part = round((amount - integer_part) * 100)

    if integer_part == 0:
        words = ["нуль"]
    else:
        words = []
        order = 0
        while integer_part > 0 and order < len(ORDERS):
            n = integer_part % 1000
            if n > 0:
                group_words = convert_triplet(n, ORDERS[order][1])
                group_name = ORDERS[order][0][get_plural_form(n)]
                words = group_words + [group_name] + words
            integer_part //= 1000
            order += 1

    kop = f"{fractional_part:02d}"
    kop_words = f"{kop} копiйок."

    return f"{' '.join(words)}, {kop_words}"