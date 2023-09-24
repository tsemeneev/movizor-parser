@echo off

REM Создание виртуального окружения
python -m venv venv

REM Активация виртуального окружения
call venv\Scripts\activate.bat

REM Установка зависимостей
pip install -r requirements.txt

REM Запуск файла web.py
python ctk.py