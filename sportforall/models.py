# sportforall/models.py

import uuid # Добавьте импорт uuid
import datetime # Для работы с датами (если нужно в моделях)
from typing import Dict, Any, List, Optional # Используем Optional для None
# import sys # Не нужен здесь, если constants импортируется напрямую
# import re # Если используется в моделях

# Импорт модуля constants (убедитесь, что файл constants.py существует)
# from sportforall import constants # Закомментируем, если не используется напрямую в моделях


class Field:
    # Класс для определения структуры поля (необязателен, но полезен)
    def __init__(self, name: str, field_type: str = "text", default_value: Any = ""):
        self.name = name
        self.field_type = field_type # Например: "text", "number", "date", "boolean"
        self.default_value = default_value

    def to_dict(self) -> dict:
         return {
              'name': self.name,
              'type': self.field_type,
              'default': self.default_value
         }

    @classmethod
    def from_dict(cls, data_dict: dict) -> 'Field':
         return cls(
              name=data_dict.get('name', ''),
              field_type=data_dict.get('type', 'text'),
              default_value=data_dict.get('default', '')
         )

    def __repr__(self) -> str:
         return f"Field(name='{self.name}', type='{self.field_type}')"


class Item:
    # Класс для представления товара или услуги
    def __init__(self, name: str = "", dk: str = "", quantity: float = 0.0, price: float = 0.0, id: str | None = None):
        self.id = id if id else str(uuid.uuid4()) # Генеруємо ID, якщо не задано
        self.name = name
        self.dk = dk # ДК 021:2015
        # Храним как float
        self.quantity = quantity
        self.price = price

    def get_total_sum(self) -> float:
        """Розраховує загальну суму для цього товару."""
        try:
             # Убедимся, что работаем с float, обрабатываем None
             q = float(self.quantity) if self.quantity is not None else 0.0
             p = float(self.price) if self.price is not None else 0.0
             return round(q * p, 2) # Округляем до 2 знаков после запятой
        except (TypeError, ValueError):
            # TODO: Логировать ошибку, если quantity или price не являются числами
            # print(f"Помилка розрахунку суми для товару '{self.name}': Некоректні дані ({self.quantity}, {self.price})")
            return 0.0

    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'name': self.name,
            'dk': self.dk,
            'quantity': self.quantity,
            'price': self.price,
            # total_sum можно не сохранять, т.к. она вычисляется
        }

    @classmethod
    def from_dict(cls, data_dict: dict) -> 'Item':
        """Створює об'єкт Item зі словника."""
        # Десериализуем числовые значения, обрабатывая возможные ошибки и пустые строки
        # Убедимся, что получаем float или 0.0
        def parse_float(value):
             if value is None or str(value).strip() == "":
                  return 0.0
             try:
                  # Заменяем запятую на точку для корректного преобразования
                  return float(str(value).replace(",", "."))
             except (ValueError, TypeError):
                  # TODO: Логировать ошибку десериализации конкретного поля Item
                  # print(f"Помилка десериалізації числового поля: Некоректне значення '{value}'")
                  return 0.0 # Default to 0 on error

        item = cls(
            name=data_dict.get('name', ''),
            dk=data_dict.get('dk', ''),
            quantity=parse_float(data_dict.get('quantity')),
            price=parse_float(data_dict.get('price')),
            id=data_dict.get('id')
        )
        return item

    def __repr__(self) -> str:
        return f"Item(id='{self.id}', name='{self.name}', quantity={self.quantity}, price={self.price}, total_sum={self.get_total_sum()})"


class Contract:
    def __init__(self, id: str, name: str, fields: Dict[str, Any] = None, items: List[Item] = None, template_path: str = ""):
        self.id = id
        self.name = name
        # Используем словарь для хранения динамических полей {placeholder_name: value}
        self.fields: Dict[str, Any] = fields if fields is not None else {}
        # Список для хранения товаров договора (Item objects)
        self.items: List[Item] = items if items is not None else []
        self.template_path: str = template_path # Путь к шаблону, специфичный для этого договора

    # Методы update_field, add_item, remove_item, find_item_by_id
    # могут остаться здесь или быть перемещены в AppData для централизованного управления

    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'name': self.name,
            # Сохраняем поля как есть (словарь, значения могут быть любых сериализуемых типов)
            'fields': self.fields,
            # Сериализуем список Item объектов
            'items': [item.to_dict() for item in self.items],
            'template_path': self.template_path
        }

    @classmethod
    def from_dict(cls, data_dict: dict) -> 'Contract':
        """Створює об'єкт Contract зі словника."""
        # Убедимся, что name и id существуют, предоставляем значения по умолчанию
        contract = cls(
            id=data_dict.get('id', str(uuid.uuid4())), # Генеруємо новий ID, якщо старого немає
            name=data_dict.get('name', 'Безымянный договор'),
            fields=data_dict.get('fields', {}), # Десериализуем поля
            template_path=data_dict.get('template_path', '')
        )

        # Десериализуем список Item объектов
        contract.items = []
        items_data = data_dict.get('items', [])
        if isinstance(items_data, list):
             for i_dict in items_data:
                 try:
                     item = Item.from_dict(i_dict)
                     contract.items.append(item)
                 except Exception as e:
                     # TODO: Логировать ошибку десериализации Item
                     # print(f"Помилка десериалізації товару з даних договору '{contract.name}': {e} (дані: {i_dict})")
                     pass # Пропускаем этот Item

        return contract

    def __repr__(self) -> str:
        # Улучшенное представление объекта для отладки
        return f"Contract(id='{self.id}', name='{self.name}', fields_count={len(self.fields)}, items_count={len(self.items)})"


class Event:
    def __init__(self, id: str, name: str, contracts: Dict[str, Contract] = None):
        self.id = id
        self.name = name
        # Используем словарь: id договора -> объект договора
        self.contracts: Dict[str, Contract] = contracts if contracts is not None else {}

    def add_contract(self, contract: Contract):
        """Додає договір до словника договорів заходу по ID."""
        if isinstance(contract, Contract): # Убедимся, что это объект Contract
             self.contracts[contract.id] = contract
             # print(f"Додано договір: {contract.name} (ID: {contract.id}) до заходу {self.name}") # Для дебагу
        # else:
             # print(f"Помилка: Спроба додати не-Contract об'єкт до заходу {self.name}.")


    def remove_contract(self, contract_id: str):
        """Видаляє договір за його ID зі словника договорів заходу."""
        if contract_id in self.contracts:
            # print(f"Видалено договір з ID: {contract_id} із заходу {self.name}") # Для дебагу
            del self.contracts[contract_id]
        # else:
            # print(f"Договір з ID {contract_id} не знайдено в заході {self.name} для видалення.")


    def get_contract(self, contract_id: str) -> Contract | None:
         """Повертає об'єкт договору в мероприятии по ID."""
         return self.contracts.get(contract_id)


    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'name': self.name,
            # Сериализуем словарь Contract объектов
            # Важно! Сериализуем ПОЛНЫЙ объект договора здесь, если он привязан к мероприятию
            'contracts': {c_id: contract.to_dict() for c_id, contract in self.contracts.items()}
        }

    @classmethod
    def from_dict_basic(cls, data_dict: dict) -> 'Event':
         """
         Створює об'єкт Event зі словника, НЕ десериалізуючи повністю контракти.
         Використовується при початковому завантаженні AppData для уникнення циклічних залежностей.
         Контракти будуть прив'язані пізніше по ID.
         """
         # Убедимся, что name и id существуют, предоставляем значения по умолчанию
         event = cls(name=data_dict.get('name', 'Безымянное мероприятие'), id=data_dict.get('id', str(uuid.uuid4())))
         # Не десериализуем контракты здесь! Они будут связаны позже по ID в AppData.load_data
         # event.contracts = {} # Уже инициализируется в __init__
         return event

    # Оригинальный from_dict, который десериализует контракты (менее желателен для AppData.load_data)
    # @classmethod
    # def from_dict(cls, data_dict: dict) -> 'Event':
    #      """Створює об'єкт Event зі словника, десериалізуючи контракти."""
    #      event = cls(name=data_dict.get('name', 'Безымянное мероприятие'), id=data_dict.get('id', str(uuid.uuid4())))
    #
    #      contracts_data = data_dict.get('contracts', {})
    #      event.contracts = {}
    #      if isinstance(contracts_data, dict):
    #           for c_id, c_dict in contracts_data.items():
    #                try:
    #                     contract = Contract.from_dict(c_dict)
    #                     # Проверяем, что ID в словаре совпадает с ID в данных контракта
    #                     if c_id != contract.id:
    #                          print(f"Попередження: Неспівпадіння ID контракту при десериалізації заходу. Використовуємо ID з даних: {contract.id} (очікувано {c_id}).")
    #                     event.contracts[contract.id] = contract # Всегда используем ID из данных контракта как ключ
    #                except Exception as e:
    #                     print(f"Помилка десериалізації контракту із заходу: {e} (дані: {c_dict})")
    #                     pass
    #
    #      return event


    def __repr__(self) -> str:
        # Улучшенное представление объекта для отладки
        return f"Event(id='{self.id}', name='{self.name}', contracts_count={len(self.contracts)})"