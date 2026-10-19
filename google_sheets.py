"""Модуль для CRUD-операций с Google Таблицами через Google Sheets API.

Подключение выполняется от имени сервисного аккаунта (JSON-ключ).
Оформлен как класс :class:`GoogleSheetsClient` — его можно импортировать
в любое приложение и вызывать методы напрямую.

Пример:
    from google_sheets import GoogleSheetsClient

    gs = GoogleSheetsClient()
    gs.create_row(["Иван", 30])
    print(gs.read_all())
"""

import os
import re

import gspread
from dotenv import load_dotenv

load_dotenv()

CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SPREADSHEET_ID = os.getenv("GOOGLE_SPREADSHEET_ID", "")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "module_9_lesson_2")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME", "Лист1")

# Разрешения: работа с таблицами и поиск таблиц по названию (Drive API)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _hex_to_rgb(color: str) -> dict:
    """Переводит цвет из HEX («1F4E79») в формат Google API (RGB 0..1)."""
    color = str(color).strip().lstrip("#")
    if len(color) == 3:
        color = "".join(ch * 2 for ch in color)
    if len(color) != 6:
        raise ValueError(f"Некорректный HEX-цвет: {color!r}")
    return {
        "red": int(color[0:2], 16) / 255,
        "green": int(color[2:4], 16) / 255,
        "blue": int(color[4:6], 16) / 255,
    }


def _col_to_index(column: str) -> int:
    """Переводит букву колонки ("A", "F") в 0-индекс ("A" -> 0, "F" -> 5)."""
    idx = 0
    for ch in str(column).strip().upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Некорректное имя колонки: {column!r}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _a1_to_grid_range(range_name: str, sheet_id: int) -> dict:
    """Переводит нотацию A1 ("B2", "A1:F3") в GridRange для Google API.

    GridRange указывает границы диапазона 0-индексами: начало входит,
    конец НЕ входит (для "A1:F3" последняя строка/колонка — E2).

    Аргументы:
        range_name: строка в нотации A1 (содержимое ячейка или диапазон).
        sheet_id: числовой ID листа внутри таблицы.

    Возвращает:
        Словарь вида {sheetId, startRowIndex, endRowIndex,
        startColumnIndex, endColumnIndex} для мультиметодов формата.
    """
    range_name = str(range_name).strip().upper()
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", range_name)
    if not m:
        raise ValueError(f"Поддерживаются диапазоны вида 'A1' или 'A1:F3', получено: {range_name!r}")
    c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
    if c2 is not None:
        # строки: номер строки 1-индексный, а endRowIndex — эксклюзивный, поэтому max(...) как есть
        start_row = min(int(r1), int(r2)) - 1
        end_row = max(int(r1), int(r2))
        # колонки: буква даёт 0-индекс, а endColumnIndex — эксклюзивный, поэтому +1
        col_a = _col_to_index(c1)
        col_b = _col_to_index(c2) + 1
        start_col, end_col = min(col_a, col_b), max(col_a, col_b)
    else:
        start_row, end_row = int(r1) - 1, int(r1)
        start_col, end_col = _col_to_index(c1), _col_to_index(c1) + 1
    return {
        "sheetId": sheet_id,
        "startRowIndex": start_row,
        "endRowIndex": end_row,
        "startColumnIndex": start_col,
        "endColumnIndex": end_col,
    }


# Готовые числовые форматы для параметра number_format
_NUMBER_FORMATS = {
    "date": {"type": "DATE", "pattern": "dd.mm.yyyy"},
    "money": {"type": "NUMBER", "pattern": "#,##0.00 \"₽\""},
}

_BORDER_SIDES = ("top", "bottom", "left", "right")

_H_ALIGN = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "justify": "JUSTIFIED"}
_V_ALIGN = {"top": "TOP", "middle": "MIDDLE", "bottom": "BOTTOM"}


def _build_cell_format(style: dict) -> dict:
    """Собирает JSON значения userEnteredFormat из стилей format_cells().

    Возвращает словарь только из переданных свойств — Google API умеет
    частичные обновления (мы потом перечисляем их в fields).
    """
    fmt: dict = {}

    text = {}
    if "bold" in style:
        text["bold"] = bool(style["bold"])
    if "italic" in style:
        text["italic"] = bool(style["italic"])
    if "font_size" in style:
        text["fontSize"] = int(style["font_size"])
    if "font_color" in style:
        text["foregroundColor"] = _hex_to_rgb(style["font_color"])
    if text:
        fmt["textFormat"] = text

    if "background" in style:
        fmt["backgroundColor"] = _hex_to_rgb(style["background"])

    if "number_format" in style:
        nf = style["number_format"]
        if nf is None:
            pass
        elif nf in _NUMBER_FORMATS:
            fmt["numberFormat"] = _NUMBER_FORMATS[nf].copy()
        else:
            fmt["numberFormat"] = {"type": "NUMBER", "pattern": str(nf)}

    if "halign" in style and style["halign"] is not None:
        try:
            fmt["horizontalAlignment"] = _H_ALIGN[str(style["halign"]).lower()]
        except KeyError:
            raise ValueError(f"halign: одно из {sorted(_H_ALIGN)}, получено {style['halign']!r}")

    if "valign" in style and style["valign"] is not None:
        try:
            fmt["verticalAlignment"] = _V_ALIGN[str(style["valign"]).lower()]
        except KeyError:
            raise ValueError(f"valign: одно из {sorted(_V_ALIGN)}, получено {style['valign']!r}")

    if style.get("wrap"):
        fmt["wrapStrategy"] = "WRAP"

    border_color = style.get("border_color")
    border = style.get("border")
    per_side = {side: style.get(f"border_{side}") for side in _BORDER_SIDES if style.get(f"border_{side}")}
    if border or per_side:
        color = _hex_to_rgb(border_color) if border_color else {
            "red": 0.749, "green": 0.749, "blue": 0.749,
        }
        sides = {side: border for side in _BORDER_SIDES}
        sides.update(per_side)
        fmt["borders"] = {
            side: {"style": str(val).upper(), "color": color}
            for side, val in sides.items()
            if val
        }

    return fmt


class GoogleSheetsClient:
    """Обёртка над Google Sheets API: открывает таблицу и выполняет CRUD.

    Логика двух режимов открытия таблицы:
    - если передан/задан ``spreadsheet_id`` — открывает напрямую по ID
      (надёжно: ID виден в URL таблицы, ищется только через Sheets API);
    - иначе — ищет таблицу по названию через Drive API.

    Атрибуты:
        spreadsheet_id: ID таблицы из URL (или пустая строка, если не задан).
        spreadsheet_name: название таблицы (для поиска по названию).
        worksheet_name: название листа внутри таблицы.
        spreadsheet: объект gspread.Spreadsheet — вся таблица.
        sheet: объект gspread.Worksheet — активный лист, с ним и работаем.
    """

    def __init__(
        self,
        spreadsheet_id: str = SPREADSHEET_ID,
        spreadsheet_name: str = SPREADSHEET_NAME,
        worksheet_name: str = WORKSHEET_NAME,
        credentials_file: str = CREDENTIALS_FILE,
    ):
        """Подключается к Google API и открывает лист для работы.

        Аргументы:
            spreadsheet_id: ID таблицы из URL (приоритетнее названия).
            spreadsheet_name: название таблицы (используется без ID).
            worksheet_name: название листа; если его нет — будет создан.
            credentials_file: путь к JSON-ключу сервисного аккаунта.

        Вызывает исключение:
            gspread.SpreadsheetNotFound: таблица не найдена ни по ID, ни по названию
                (или не выдан доступ сервисному аккаунту).
        """
        self.credentials_file = credentials_file
        self.spreadsheet_id = spreadsheet_id
        self.spreadsheet_name = spreadsheet_name
        self.worksheet_name = worksheet_name

        self.client = self._authorize()
        self.spreadsheet = self._open_spreadsheet()
        self.sheet = self._open_worksheet()

    @classmethod
    def list_available(cls, credentials_file: str = CREDENTIALS_FILE) -> list[str]:
        """Возвращает названия всех таблиц, доступных сервисному аккаунту.

        Полезно, чтобы узнать точные названия таблиц и их ID, к которым
        аккаунт имеет доступ (если у него есть хоть одна таблица).

        Аргументы:
            credentials_file: путь к JSON-ключу сервисного аккаунта.

        Возвращает:
            Список строк с названиями таблиц (может быть пустым).
        """
        gc = gspread.service_account(filename=credentials_file, scopes=SCOPES)
        return [file.title for file in gc.openall()]

    def _authorize(self) -> gspread.Client:
        """Создаёт клиент Google API из JSON-ключа сервисного аккаунта."""
        return gspread.service_account(filename=self.credentials_file, scopes=SCOPES)

    def _open_spreadsheet(self) -> gspread.Spreadsheet:
        """Открывает таблицу: сначала по ID (если задан), иначе по названию."""
        if self.spreadsheet_id:
            return self.client.open_by_key(self.spreadsheet_id)
        return self.client.open(self.spreadsheet_name)

    def _open_worksheet(self) -> gspread.Worksheet:
        """Открывает лист по названию; если листа нет — создаёт новый."""
        try:
            return self.spreadsheet.worksheet(self.worksheet_name)
        except gspread.WorksheetNotFound:
            return self.spreadsheet.add_worksheet(title=self.worksheet_name, rows=100, cols=26)

    # CREATE
    def create_row(self, values: list) -> list:
        """Добавляет одну строку в конец листа.

        Аргументы:
            values: список значений ячеек строки, напр. ["Иван", 30].

        Возвращает:
            Список ячеек, которые были реально обновлены.
        """
        return self.sheet.append_row(values, value_input_option="USER_ENTERED")

    def create_rows(self, matrix: list[list]) -> None:
        """Добавляет сразу несколько строк в конец листа.

        Аргументы:
            matrix: список строк, каждая строка — список значений ячеек.
        """
        self.sheet.append_rows(matrix, value_input_option="USER_ENTERED")

    # READ
    def read_all(self) -> list[list]:
        """Читает все ячейки листа.

        Возвращает:
            Список строк: каждая строка — список строковых значений ячеек,
            начиная с A1 и до последней заполненной.
        """
        return self.sheet.get_all_values()

    def read_records(self) -> list[dict]:
        """Читает данные как список словарей (первая строка — заголовки).

        Возвращает:
            [{«Имя»: «Иван», «Возраст»: 30}, ...] по строкам под заголовками.
        """
        return self.sheet.get_all_records()

    def read_range(self, range_name: str) -> list[list]:
        """Читает конкретный диапазон, напр. "A1:C5".

        Аргументы:
            range_name: диапазон в нотации A1 или A1:type.

        Возвращает:
            Список строк диапазона (только заполненные ячейки).
        """
        return self.sheet.get(range_name)

    def read_cell(self, row: int, col: int) -> str:
        """Читает значение одной ячейки по номерам строки и столбца (с 1).

        Аргументы:
            row: номер строки (1 — самая первая).
            col: номер столбца (1 — столбец A).

        Возвращает:
            Значение ячейки или None, если она пустая.
        """
        return self.sheet.cell(row, col).value

    def find(self, text: str) -> gspread.cell.Cell | None:
        """Ищет первую ячейку с данным текстом на листе.

        Аргументы:
            text: текст для поиска.

        Возвращает:
            Объект gspread.Cell (со свойствами row/col/value) или None.
        """
        return self.sheet.find(text)

    # UPDATE
    def update_cell(self, row: int, col: int, value) -> None:
        """Обновляет (перезаписывает) одну ячейку.

        Аргументы:
            row: номер строки (с 1).
            col: номер столбца (с 1, A = 1).
            value: новое значение ячейки (или формула, например "=SUM(A1:A2)").
        """
        self.sheet.update_cell(row, col, value)

    def update_range(self, range_name: str, matrix: list[list]) -> None:
        """Перезаписывает диапазон данными из матрицы.

        Значения вставляются в режиме USER_ENTERED: Google сам
        распознаёт числа, даты и формулы (например "=SUM(A1:A2)").

        Аргументы:
            range_name: диапазон записи, напр. "A1:B2".
            matrix: список строк со значениями,
                напр. [["Имя", "Возраст"], ["Иван", 30]].
        """
        self.sheet.update(
            values=matrix,
            range_name=range_name,
            value_input_option="USER_ENTERED",
        )

    # DELETE
    def delete_row(self, row_index: int) -> int:
        """Удаляет строку по её индексу (0 — самая первая строка).

        Аргументы:
            row_index: индекс строки (нумерация с нуля, как в Python).

        Возвращает:
            Номер удалённой строки (1 — для первой).
        """
        self.sheet.delete_rows(row_index + 1)
        return row_index + 1

    def delete_last_row(self) -> int | None:
        """Удаляет последнюю заполненную строку (если таблица не пуста).

        Возвращает:
            Номер удалённой строки или None, если лист пуст.
        """
        last_row = len(self.sheet.get_all_values())
        if last_row:
            self.sheet.delete_rows(last_row)
            return last_row
        return None

    def clear(self) -> None:
        """Очищает весь лист от данных (сам лист при этом остаётся)."""
        self.sheet.clear()

    # ФОРМАТИРОВАНИЕ (ОФОРМЛЕНИЕ ЛИСТА)
    def set_worksheet(self, title: str) -> gspread.Worksheet:
        """Переключает активный лист на лист с данным названием.

        Если листа с таким названием нет — он создаётся (пустой,
        200 строк x 26 колонок). Все методы этого класса после вызова
        работают именно с этим листом.

        Аргументы:
            title: название листа, напр. "Отчёт_19.10.2026_09-30".

        Возвращает:
            Активный лист (gspread.Worksheet).
        """
        try:
            self.sheet = self.spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            self.sheet = self.spreadsheet.add_worksheet(title=title, rows=200, cols=26)
        return self.sheet

    def _run_batch(self, requests: list[dict]) -> None:
        """Отправляет список частичных запросов в API одним вызовом."""
        self.spreadsheet.batch_update({"requests": requests})

    def format_cells(self, range_name: str, **style) -> None:
        """Форматирует диапазон ячеек (шрифт, фон, рамки, выравнивание).

        Меняются только переданные свойства — остальное оформление
        ячеек сохраняется. Стиль задаётся именованными аргументами:
            bold=True/False          — жирный шрифт
            italic=True/False        — курсив
            font_size=12             — размер шрифта
            font_color="#FF0000"     — цвет шрифта (HEX)
            background="#D9E2F1"     — цвет заливки (HEX)
            border="SOLID"           — стиль рамки со всех сторон:
                                       SOLID / SOLID_MEDIUM / SOLID_THICK /
                                       DASHED / DOTTED / DOUBLE
            border_color="#BFBFBF"   — цвет рамки (HEX)
            halign="center"          — выравнивание: left/center/right
            valign="middle"          — по вертикали: top/middle/bottom
            wrap=True                — переносить длинный текст по строкам
            number_format="date"     — формат числа: "date" (ДД.ММ.ГГГГ)
                                       или "money" (#,##0.00 ₽)
                                       или своя маска, напр. "0.0%"

        Аргументы:
            range_name: диапазон, напр. "A1:F1" или "B6:B12".
            **style: свойства оформления (см. выше).
        """
        self.batch_format([(range_name, style)])

    def batch_format(self, specs: list[tuple[str, dict]]) -> None:
        """Форматирует несколько диапазонов за один запрос к API.

        Удобно, когда диапазонов много (например, «зебра» из чередующихся
        строк): вместо десятка вызовов format_cells выполняется один.

        Аргументы:
            specs: список пар (диапазон, стиль), где стиль — такие же
                именованные свойства, как у format_cells().
        """
        requests = []
        for range_name, style in specs:
            fmt = _build_cell_format(style)
            user_fmt = {"userEnteredFormat": fmt}
            fields = ",".join(f"userEnteredFormat.{key}" for key in fmt)
            requests.append({
                "repeatCell": {
                    "range": _a1_to_grid_range(range_name, self.sheet.id),
                    "cell": user_fmt,
                    "fields": fields,
                }
            })
        self._run_batch(requests)

    def merge_range(self, range_name: str) -> None:
        """Объединяет ячейки диапазона в одну (например, заголовок A1:F1)."""
        body = [{
            "mergeCells": {
                "range": _a1_to_grid_range(range_name, self.sheet.id),
                "mergeType": "MERGE_ALL",
            }
        }]
        self._run_batch(body)

    def set_column_width(self, column: str, width: int) -> None:
        """Задаёт ширину колонки в пикселях, напр. set_column_width("A", 40)."""
        idx = _col_to_index(column)
        body = [{
            "updateDimensionProperties": {
                "range": {
                    "sheetId": self.sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": idx,
                    "endIndex": idx + 1,
                },
                "properties": {"pixelSize": width},
                "fields": "pixelSize",
            }
        }]
        self._run_batch(body)

    def autofit_columns(self, first: str = "A", last: str = "F") -> None:
        """Автоматически подбирает ширину колонок под содержимое."""
        body = [{
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": self.sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": _col_to_index(first),
                    "endIndex": _col_to_index(last) + 1,
                }
            }
        }]
        self._run_batch(body)

    def set_row_height(self, row: int, height: int) -> None:
        """Задаёт высоту строки в пикселях, напр. set_row_height(1, 32)."""
        body = [{
            "updateDimensionProperties": {
                "range": {
                    "sheetId": self.sheet.id,
                    "dimension": "ROWS",
                    "startIndex": row - 1,
                    "endIndex": row,
                },
                "properties": {"pixelSize": height},
                "fields": "pixelSize",
            }
        }]
        self._run_batch(body)

    def freeze_rows(self, rows: int) -> None:
        """Закрепляет верхние строки: при прокрутке они не уезжают.

        Аргументы:
            rows: сколько первых строк закрепить (0 — ничего).
        """
        body = [{
            "updateSheetProperties": {
                "properties": {
                    "sheetId": self.sheet.id,
                    "gridProperties": {"frozenRowCount": rows},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }]
        self._run_batch(body)

    def hide_gridlines(self) -> None:
        """Скрывает линии сетки (лист выглядит как обычный документ)."""
        body = [{
            "updateSheetProperties": {
                "properties": {
                    "sheetId": self.sheet.id,
                    "gridProperties": {"hideGridlines": True},
                },
                "fields": "gridProperties.hideGridlines",
            }
        }]
        self._run_batch(body)

    def remove_worksheet(self, title: str) -> None:
        """Удаляет лист по названию (иначе ошибку WorksheetNotFound)."""
        ws = self.spreadsheet.worksheet(title)
        self.spreadsheet.del_worksheet(ws)


if __name__ == "__main__":
    gs = GoogleSheetsClient()

    source = f"ID: {gs.spreadsheet_id}" if gs.spreadsheet_id else f"название: {gs.spreadsheet_name}"
    print(f"Таблица ({source})")
    print(f"Лист:    {gs.worksheet_name}")

    data = gs.read_all()
    print(f"Всего строк: {len(data)}")
    print("-" * 50)
    for row in data:
        print(row)
