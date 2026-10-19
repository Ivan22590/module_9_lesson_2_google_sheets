# CRUD-модуль для Google Таблиц (Google Sheets API)

Модуль `google_sheets.py` даёт класс `GoogleSheetsClient` для работы с
Google Таблицами через API от имени сервисного аккаунта. Его можно
подключать в любое приложение и вызывать методы напрямую.

## Состав проекта

| Файл | Назначение |
|---|---|
| `google_sheets.py` | Класс `GoogleSheetsClient` (CRUD + форматирование) + точка входа |
| `report_generator.py` | Tkinter-приложение «Генератор отчётов» |
| `.env` | Реальные переменные окружения (не коммитится) |
| `.env.example` | Шаблон переменных окружения |
| `requirements.txt` | Зависимости проекта |
| `excel-factory-512109-668e007b2183.json` | Ключ сервисного аккаунта (не коммитится) |

## Быстрый старт

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env   # затем заполните реальные значения
```

В `.env` укажите ключ сервисного аккаунта, ID таблицы и лист:

```
GOOGLE_CREDENTIALS_FILE=service-account-key.json

# ID таблицы из URL (docs.google.com/spreadsheets/d/<ID>/edit) — предпочтительный способ
GOOGLE_SPREADSHEET_ID=

# Если ID не задан — таблица ищется по названию
SPREADSHEET_NAME=my-spreadsheet
WORKSHEET_NAME=Sheet1
```

### Два способа открыть таблицу

| Способ | Как работает | Когда использовать |
|---|---|---|
| По ID (`open_by_key`) | Обращается напрямую по ID, только Sheets API | **Предпочтительно**: работает быстро и надёжно |
| По названию (`open`) | Ищет через Drive API таблицу с таким именем | Когда ID неизвестен; при одноимённых таблицах выдаст неожиданный результат |

### ВАЖНО: доступ к таблице

Сервисный аккаунт видит только те таблицы, к которым ему выдан доступ.
Откройте нужную таблицу в Google Sheets → «Поделиться» → вставьте email
сервисного аккаунта (поле `client_email` из JSON-ключа) в роли «Редактор».

## Проверка (точка входа)

```bash
python google_sheets.py
```

Скрипт прочитает все ячейки листа и напечатает их в консоль.

## Приложение «Генератор отчётов» (Tkinter)

Окно с полями: период (от/до), тип отчёта, подразделение. По кнопке
«Сформировать и выгрузить» приложение генерирует случайные данные
за период и записывает их в Google Таблицу на **новый лист**
`Отчёт_<тип>_<дата>_<время>` в виде «документа»: объединённый заголовок,
шапка с заливкой, рамки, чередование строк, денежный формат ₽,
строка ИТОГО, авторазмер колонок, скрытая сетка, закреплённая шапка.

```bash
python report_generator.py
```

Требует установленного tkinter: на Windows он входит в стандартный
инсталлятор Python, в Debian/Ubuntu — `sudo apt install python3-tk`.

Логика отчёта отделена от окна и работает без GUI:

```python
from datetime import date
from report_generator import write_report
from google_sheets import GoogleSheetsClient

gs = GoogleSheetsClient()
title = write_report(gs, date(2026, 10, 1), date(2026, 10, 31), "Продажи", "Маркетинг")
print("Создан лист:", title)
```

## Использование в другом приложении

```python
from google_sheets import GoogleSheetsClient

# Способ 1: по ID из URL таблицы (рекомендуется)
gs = GoogleSheetsClient(spreadsheet_id="1AbCdEf...", worksheet_name="Sheet1")

# Способ 2: по названию (фолбэк, если ID не задан)
gs = GoogleSheetsClient(spreadsheet_name="my-spreadsheet", worksheet_name="Sheet1")

# Список таблиц, доступных сервисному аккаунту (названия)
print(GoogleSheetsClient.list_available())

# CREATE
gs.create_row(["Иван", 30])

# READ
print(gs.read_all())            # все ячейки списком строк
print(gs.read_records())        # список словарей (первая строка = заголовки)
print(gs.read_range("A1:B2"))   # конкретный диапазон
print(gs.read_cell(1, 1))       # ячейка A1
found = gs.find("Иван")         # поиск ячейки

# UPDATE
gs.update_cell(2, 1, "Пётр")            # A2
gs.update_range("A1:B1", [["Имя", "Возраст"]])

# DELETE
gs.delete_last_row()
gs.clear()                      # очистить весь лист
```

## Методы класса

| Метод | Описание |
|---|---|
| `create_row(values)` | Добавить строку в конец листа |
| `create_rows(matrix)` | Добавить несколько строк |
| `read_all()` | Все значения листа (список строк) |
| `read_records()` | Значения как список словарей |
| `read_range(range_name)` | Чтение диапазона, напр. `"A1:C5"` |
| `read_cell(row, col)` | Значение одной ячейки |
| `find(text)` | Найти первую ячейку с текстом |
| `update_cell(row, col, value)` | Обновить ячейку |
| `update_range(range_name, matrix)` | Обновить диапазон матрицей |
| `delete_row(row_index)` | Удалить строку (0-индексация); возвращает её номер |
| `delete_last_row()` | Удалить последнюю строку; возвращает её номер или `None` |
| `clear()` | Очистить лист |
| `set_worksheet(title)` | Переключить активный лист (создать, если нет) |
| `format_cells(range, **style)` | Оформление: шрифт, фон, рамки, выравнивание, формат чисел |
| `batch_format([(range, style), …])` | Форматирование нескольких диапазонов за один запрос |
| `merge_range(range)` | Объединить ячейки (заголовки документа) |
| `set_column_width(col, px)` / `set_row_height(row, px)` | Точные размеры |
| `autofit_columns(first, last)` | Автоширина колонок под содержимое |
| `freeze_rows(n)` / `hide_gridlines()` | Закрепление шапки и «бумажный» вид |
| `remove_worksheet(title)` | Удалить лист |
