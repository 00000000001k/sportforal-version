# sportforall/gui/__init__.py

# Импортируем основные классы GUI, чтобы их можно было импортировать напрямую из пакета GUI
# Теперь главный класс называется MainApp и находится в main_app.py
from .main_app import MainApp
from .custom_widgets import CustomEntry # SafeCTk нам не нужен, если не используем окно пароля
# Импортируем другие нужные классы View, если их планируется использовать напрямую из gui пакета
from .event_contract_views import EventContractViews
from .contract_details_view import ContractDetailsView


# Не импортируем здесь вспомогательные функции или модули,
# которые используются только внутри других модулей View (например, error_handling).