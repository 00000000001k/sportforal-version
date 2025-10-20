from num2words import num2words

def number_to_ukrainian_text(amount):
    try:
        amount = float(str(amount).replace(",", "."))
    except:
        return "Нуль гривень, 00 копійок."

    integer_part = int(amount)
    fractional_part = int(round((amount - integer_part) * 100))

    # Текстова частина суми
    words = num2words(integer_part, lang='uk')

    # Узгодження "гривня"
    def plural(n, forms):
        if n % 10 == 1 and n % 100 != 11:
            return forms[0]
        elif 2 <= n % 10 <= 4 and not (12 <= n % 100 <= 14):
            return forms[1]
        else:
            return forms[2]

    hryvnia_word = plural(integer_part, ["гривня", "гривні", "гривень"])
    kopiyky_word = plural(fractional_part, ["копійка", "копійки", "копійок"])

    return f"{words} {hryvnia_word}, {fractional_part:02d} {kopiyky_word}."
