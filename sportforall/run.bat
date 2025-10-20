@echo off
echo Запуск программы Генератор договорів...
python run.py
if %ERRORLEVEL% NEQ 0 (
    echo Программа завершилась с ошибкой!
    echo Детали ошибки можно найти в файле error.txt
    pause
)