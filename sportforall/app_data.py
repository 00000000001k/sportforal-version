# sportforall/app_data.py

import json
import os
import uuid # Для генерации уникальных ID
import sys # Для использования sys.exc_info
from typing import Dict, Any, List

# Импортируем модели данных (теперь они в отдельном файле models.py)
from sportforall.models import Contract, Event, Field, Item # Убедитесь, что все модели импортированы
from sportforall import error_handling # Модуль для логирования ошибок


# Хранение данных приложения (договоры, мероприятия, настройки)
class AppData:
    def __init__(self, filename: str = "app_data.json"):
        self.filename = filename
        # Используем словари для быстрого доступа по ID
        self.contracts: Dict[str, Contract] = {} # id -> Contract object
        self.events: Dict[str, Event] = {}       # id -> Event object
        # Глобальные настройки
        self.template_path: str = ""
        self.output_dir: str = ""
        self.contract_default_template_name: str = "Новый договор" # Название для новых договоров
        self.event_default_name: str = "Новое мероприятие" # Название для новых мероприятий

        # TODO: Добавить загрузку и сохранение списка стандартных полей для договоров
        # self.standard_fields: Dict[str, Dict[str, Any]] = {} # Пример: {"поле1": {"type": "text", "default": ""}}


    def load_data(self):
        """Загружает данные из JSON файла."""
        if os.path.exists(self.filename):
            print(f"DEBUG APP_DATA: Loading data from {self.filename}")
            try:
                with open(self.filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                    # Загрузка настроек
                    settings = data.get("settings", {})
                    self.template_path = settings.get("template_path", "")
                    self.output_dir = settings.get("output_dir", "")
                    self.contract_default_template_name = settings.get("contract_default_template_name", "Новый договор")
                    self.event_default_name = settings.get("event_default_name", "Новое мероприятие")

                    # Загрузка договоров
                    self.contracts = {}
                    contracts_data = data.get("contracts", {})
                    if isinstance(contracts_data, dict): # Убедимся, что это словарь
                         for contract_id, contract_data in contracts_data.items():
                            try:
                                contract = Contract.from_dict(contract_data)
                                self.contracts[contract.id] = contract
                                # print(f"DEBUG APP_DATA: Loaded contract: {contract.name} ({contract.id}) with {len(contract.fields)} fields and {len(contract.items)} items.") # For debug
                            except Exception as e:
                                print(f"ERROR APP_DATA: Failed to load contract ID {contract_id}: {e}")
                                error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Ошибка загрузки договора ID {contract_id} из данных: {e}")
                                pass # Пропускаем некорректный договор


                    # Загрузка мероприятий
                    self.events = {}
                    events_data = data.get("events", {})
                    # Загружаем мероприятия сначала без привязки договоров
                    if isinstance(events_data, dict): # Убедимся, что это словарь
                         for event_id, event_data in events_data.items():
                            try:
                                event = Event.from_dict_basic(event_data) # Используем from_dict_basic, который не грузит контракты
                                self.events[event.id] = event
                                # print(f"DEBUG APP_DATA: Loaded event: {event.name} ({event.id})") # For debug
                            except Exception as e:
                                print(f"ERROR APP_DATA: Failed to load event ID {event_id}: {e}")
                                error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Ошибка загрузки мероприятия ID {event_id} из данных: {e}")
                                pass # Пропускаем некорректное мероприятие


                    # Привязка договоров к мероприятиям (после того как все договоры и мероприятия загружены)
                    if isinstance(events_data, dict):
                         for event_id, event_data in events_data.items():
                              if event_id in self.events:
                                   event = self.events[event_id]
                                   # event.contracts = {} # Инициализируем словарь контрактов для мероприятия (уже должно быть в Event.__init__)
                                   contracts_in_event_data = event_data.get("contracts", {}) # Получаем словарь ID -> contract_data

                                   if isinstance(contracts_in_event_data, dict): # Убедимся, что это словарь
                                        for contract_id, contract_in_event_data in contracts_in_event_data.items():
                                             if contract_id in self.contracts:
                                                  # Привязываем существующий объект договора по ID
                                                  event.contracts[contract_id] = self.contracts[contract_id]
                                                  # print(f"DEBUG APP_DATA: Linked contract {self.contracts[contract_id].name} to event {event.name}") # For debug
                                             else:
                                                  # TODO: Логировать предупреждение, если привязанный договор не найден
                                                  print(f"WARNING APP_DATA: Contract ID {contract_id} not found when linking to event {event.name}. Skipping link.")
                                                  error_handling.log_error(None, None, None, level="WARNING", message=f"Договор ID {contract_id} не найден при привязке к мероприятию {event.name}.")
                                   # else:
                                        # TODO: Логировать предупреждение о некорректном формате данных контрактов в мероприятии
                                        # print(f"WARNING APP_DATA: Unexpected format for contracts data in event {event.name}: {type(contracts_in_event_data)}. Expected dict.")
                                        # error_handling.log_error(None, None, None, level="WARNING", message=f"Неожиданный формат данных контрактов в мероприятии {event.name}.")


                print("DEBUG APP_DATA: Data loaded successfully.")
            except json.JSONDecodeError as e:
                print(f"ERROR APP_DATA: JSON Decode Error loading data from {self.filename}: {e}")
                error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Ошибка декодирования JSON файла данных: {e}")
                # Начать с пустыми данными после ошибки
                self._reset_data()
            except Exception as e:
                print(f"ERROR APP_DATA: An unexpected error occurred loading data from {self.filename}: {e}")
                error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Неожиданная ошибка при загрузке данных: {e}")
                 # Начать с пустыми данными после ошибки
                self._reset_data()
        else:
            print(f"DEBUG APP_DATA: Data file not found: {self.filename}. Starting with empty data.")
            # Файл не найден, начинаем с пустыми данными
            self._reset_data()

        # print(f"DEBUG APP_DATA: Loaded {len(self.contracts)} contracts and {len(self.events)} events.")


    def save_data(self):
        """Сохраняет текущие данные в JSON файл."""
        data = {
            "settings": {
                "template_path": self.template_path,
                "output_dir": self.output_dir,
                "contract_default_template_name": self.contract_default_template_name,
                "event_default_name": self.event_default_name,
            },
            "contracts": {},
            "events": {}
        }

        # Сохранение договоров
        for contract_id, contract in self.contracts.items():
             try:
                data["contracts"][contract_id] = contract.to_dict()
             except Exception as e:
                  print(f"ERROR APP_DATA: Failed to serialize contract ID {contract_id}: {e}")
                  error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Ошибка сериализации договора ID {contract_id}: {e}")
                  pass # Пропускаем проблемный договор


        # Сохранение мероприятий и привязанных к ним договоров (полностью, как словарь id -> contract_dict)
        for event_id, event in self.events.items():
             try:
                 # Сохраняем мероприятие, включая его контракты (которые сериализуются в to_dict Event)
                 data["events"][event_id] = event.to_dict()
             except Exception as e:
                  print(f"ERROR APP_DATA: Failed to serialize event ID {event_id}: {e}")
                  error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Ошибка сериализации мероприятия ID {event_id}: {e}")
                  pass # Пропускаем проблемное мероприятие


        try:
            # Убедимся, что папка для файла данных существует
            os.makedirs(os.path.dirname(self.filename) or '.', exist_ok=True) # Создать паку, если не существует

            with open(self.filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False) # ensure_ascii=False для сохранения кириллицы

            print(f"DEBUG APP_DATA: Data successfully saved to {self.filename}")
        except Exception as e:
            print(f"ERROR APP_DATA: An error occurred saving data to {self.filename}: {e}")
            error_handling.log_error(type(e), e, sys.exc_info()[2], message=f"Ошибка при сохранении данных в файл {self.filename}: {e}")

    def _reset_data(self):
        """Сбрасывает данные приложения к пустому состоянию."""
        self.contracts = {}
        self.events = {}
        self.template_path = ""
        self.output_dir = ""
        self.contract_default_template_name = "Новый договор"
        self.event_default_name = "Новое мероприятие"


    def create_contract(self, name: str) -> Contract:
        """Создает новый объект договора и добавляет его в список."""
        new_id = str(uuid.uuid4())
        # При создании нового договора, инициализируем пустой словарь fields
        # TODO: Инициализировать новый договор стандартными полями, если они определены
        new_contract = Contract(id=new_id, name=name, fields={}) # Инициализируем поля пустым словарем
        self.contracts[new_id] = new_contract
        # self.save_data() # Можно сохранить сразу или делать это позже
        # print(f"DEBUG APP_DATA: Created new contract: {new_contract.name} ({new_contract.id})")
        return new_contract

    def delete_contract(self, contract_id: str):
        """Удаляет договор по ID."""
        if contract_id in self.contracts:
            # Удаляем договор из всех мероприятий, к которым он привязан
            for event in self.events.values():
                if contract_id in event.contracts:
                    del event.contracts[contract_id]
                    # print(f"DEBUG APP_DATA: Removed contract {contract_id} from event {event.name}")

            # Удаляем сам договор
            del self.contracts[contract_id]
            # self.save_data() # Можно сохранить сразу или делать это позже
            # print(f"DEBUG APP_DATA: Deleted contract with ID {contract_id}")


    def get_contract(self, contract_id: str) -> Contract | None:
        """Возвращает объект договора по ID из основного словаря договоров."""
        return self.contracts.get(contract_id)

    def get_event(self, event_id: str) -> Event | None:
         """Возвращает объект мероприятия по ID."""
         return self.events.get(event_id)

    def get_event_by_name(self, event_name: str) -> Event | None:
         """Возвращает объект мероприятия по имени."""
         for event in self.events.values():
              if event.name == event_name:
                   return event
         return None

    def create_event(self, name: str) -> Event:
         """Создает новое мероприятие и добавляет его."""
         new_id = str(uuid.uuid4())
         new_event = Event(id=new_id, name=name)
         self.events[new_id] = new_event
         # self.save_data()
         print(f"DEBUG APP_DATA: Created new event: {new_event.name} ({new_event.id})")
         return new_event

    def delete_event(self, event_id: str):
         """Удаляет мероприятие по ID."""
         if event_id in self.events:
              # Удаляем мероприятие
              del self.events[event_id]
              # self.save_data()
              print(f"DEBUG APP_DATA: Deleted event with ID {event_id}")

    def add_contract_to_event(self, event_id: str, contract_id: str):
         """Привязывает существующий договор к существующему мероприятию."""
         event = self.get_event(event_id)
         contract = self.get_contract(contract_id)
         if event and contract:
              event.add_contract(contract) # Метод add_contract в Event должен принимать объект
              # self.save_data()
              print(f"DEBUG APP_DATA: Linked contract {contract.name} to event {event.name}")
         # else:
              # print(f"WARNING APP_DATA: Failed to link contract {contract_id} to event {event_id}. Event or Contract not found.")

    def remove_contract_from_event(self, event_id: str, contract_id: str):
         """Отвязывает договор от мероприятия."""
         event = self.get_event(event_id)
         if event:
              event.remove_contract(contract_id)
              # self.save_data()
              print(f"DEBUG APP_DATA: Unlinked contract {contract_id} from event {event.name}")