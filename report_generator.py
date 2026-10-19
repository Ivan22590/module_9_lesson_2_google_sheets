"""Tkinter-приложение «Генератор отчётов».

Пользователь задаёт период, тип отчёта и подразделение, приложение
генерирует случайные («симулированные») данные и записывает их в
Google Таблицу в виде красиво свёрстанного отчёта на новом листе:
объединённые заголовки, шапка таблицы с заливкой, рамки, «зебра»,
денежный формат, строка ИТОГО, закрепление шапки, авторазмер колонок.

Запуск окна на обычном компьютере:
    python report_generator.py

На headless-сервере окно не открывается, но логика без интерфейса
работает — см. функции generate_report_data() и write_report().
"""

import datetime
import queue
import random
import threading

# tkinter входит в Python, но на headless-серверах его часто нет:
# при отсутствии окно не запускается, а логика отчёта остаётся рабочей
try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except ModuleNotFoundError:
    tk = None

from google_sheets import GoogleSheetsClient

# Поля интерфейса
DEPARTMENTS = ("Продажи", "Маркетинг", "Финансы", "HR", "Логистика")

# Категории данных зависят от выбранного типа отчёта
REPORT_TYPES = {
    "Продажи": ["Ноутбуки", "Смартфоны", "Планшеты", "Мониторы", "Аксессуары"],
    "Движение средств": ["Поступление", "Выплата", "Перевод", "Комиссия", "Прочее"],
    "Командировки": ["Авиабилеты", "Отели", "Транспорт", "Суточные", "Расходы на месте"],
}

EMPLOYEES = (
    "Иванов И.", "Петров П.", "Сидорова А.", "Кузнецов Д.",
    "Смирнова Е.", "Васильев С.", "Морозова О.", "Козлов Н.",
)

DESCRIPTIONS = (
    "оплата по договору", "возврат товара", "корректировка",
    "срочная закупка", "регулярная поставка", "разовая сделка",
    "предоплата", "перерасчёт", "по счёту", "закрытие периода",
)

# Цвета оформления отчёта (HEX)
NAVY = "1F4E79"      # тёмно-синий: заголовок и шапка таблицы
BLUE_LIGHT = "D9E2F1"  # светло-голубой: подзаголовок и строка ИТОГО
ZEBRA = "EAF2FB"     # чередование строк данных
BORDER = "BFBFBF"    # цвет рамок


def parse_date(value: str) -> datetime.date:
    """Разбирает строку с датой в формате ДД.ММ.ГГГГ.

    Аргументы:
        value: строка от пользователя, напр. "01.10.2026".

    Возвращает:
        Объект datetime.date.

    Вызывает исключение:
        ValueError: с понятным русским описанием, если формат неверный.
    """
    try:
        return datetime.datetime.strptime(value.strip(), "%d.%m.%Y").date()
    except ValueError:
        raise ValueError(f"Нужна дата в формате ДД.ММ.ГГГГ, получено: {value!r}")


def generate_report_data(
    date_from: datetime.date,
    date_to: datetime.date,
    report_type: str,
    department: str,
    n_rows: int | None = None,
) -> list[list]:
    """Генерирует случайные данные отчёта за указанный период.

    Данные «симулированные»: дата внутри [date_from, date_to], случайная
    категория выбранного типа отчёта, описание, сотрудник и сумма.

    Аргументы:
        date_from: первая дата периода (входит в диапазон).
        date_to: последняя дата периода (входит).
        report_type: один из ключей REPORT_TYPES, напр. "Продажи".
        department: одно из значений DEPARTMENTS.
        n_rows: сколько строк сгенерировать; None — случайное 15..40.

    Возвращает:
        Список строк без шапки: [№, дата(ГГГГ-ММ-ДД), категория,
        описание, сотрудник, сумма(float)].

    Вызывает исключение:
        ValueError: если date_to раньше date_from или тип неизвестен.
    """
    if date_to < date_from:
        raise ValueError("Дата «до» не может быть раньше даты «от»")
    if report_type not in REPORT_TYPES:
        raise ValueError(f"Неизвестный тип отчёта: {report_type!r}")
    if department not in DEPARTMENTS:
        raise ValueError(f"Неизвестное подразделение: {department!r}")

    if n_rows is None:
        n_rows = random.randint(15, 40)

    categories = REPORT_TYPES[report_type]
    span_days = (date_to - date_from).days

    rows = []
    for num in range(1, n_rows + 1):
        day_offset = random.randint(0, span_days)
        row_date = date_from + datetime.timedelta(days=day_offset)
        amount = round(random.uniform(500, 250_000), 2)
        rows.append([
            num,
            row_date.isoformat(),
            random.choice(categories),
            random.choice(DESCRIPTIONS),
            random.choice(EMPLOYEES),
            amount,
        ])
    return rows


def build_report_matrix(
    data_rows: list[list],
    date_from: datetime.date,
    date_to: datetime.date,
    report_type: str,
    department: str,
    generated_at: datetime.datetime,
) -> tuple[list[list], dict]:
    """Собирает полную матрицу отчёта: «шапка документа» + данные + ИТОГО.

    Структура листа (колонки A..F):
        строка 1: заголовок отчёта (объединённая ячейка A1:F1)
        строка 2: период (объединённая A2:F2)
        строка 3: сформировано / подразделение (объединённая A3:F3)
        строка 4: пустой отступ
        строка 5: шапка таблицы
        строки 6..: данные
        последняя: ИТОГО с формулой суммы

    Возвращает:
        Матрицу для update_range("A1:F...") и словарь layout с номерами
        строк данных и итоговой строки (для последующего форматирования).
    """
    header = ["№", "Дата", "Категория", "Описание", "Сотрудник", "Сумма"]
    total_row_num = 6 + len(data_rows)  # строка, идущая после всех данных

    matrix = [
        [f"ОТЧЁТ: {report_type.upper()}", "", "", "", "", ""],
        [
            f"Период: {date_from:%d.%m.%Y} — {date_to:%d.%m.%Y}",
            "", "", "", "", "",
        ],
        [
            f"Сформировано: {generated_at:%d.%m.%Y %H:%M} | "
            f"Подразделение: {department}",
            "", "", "", "", "",
        ],
        ["", "", "", "", "", ""],
        header,
    ]
    matrix.extend(data_rows)
    # Формула суммы запишется числом в колонку F (mode USER_ENTERED)
    matrix.append([
        "ИТОГО:", "", "", "", "",
        f"=SUM(F6:F{total_row_num - 1})",
    ])

    layout = {
        "data_first_row": 6,
        "data_last_row": total_row_num - 1,
        "total_row": total_row_num,
        "n_columns": 6,
    }
    return matrix, layout


def format_report(gs: GoogleSheetsClient, layout: dict) -> None:
    """Оформляет только что записанный отчёт (объединения, заливки, рамки).

    Вызывается после write_report() — активный лист клиента уже отчёт.
    Оформление:
        A1:F1 — большой белый заголовок на тёмно-синем;
        A2:F2 — период: жирный, светло-голубая заливка;
        A3:F3 — сформировано: курсив;
        строка 5 — шапка: белая по синему, закрепляется;
        данные — рамки, «зебра», даты ДД.ММ.ГГГГ, суммы в ₽;
        ИТОГО — жирная, светло-голубая, двойная верхняя рамка;
        сетка скрыта, колонки B..F — авторазмер, A — узкая.
    """
    first, last = layout["data_first_row"], layout["data_last_row"]
    total = layout["total_row"]

    # Заголовок, шапка, итог — за один запрос к API
    gs.batch_format([
        ("A1:F1", {
            "bold": True, "font_size": 16, "font_color": "#FFFFFF",
            "background": NAVY, "halign": "center", "valign": "middle",
        }),
        ("A2:F2", {"bold": True, "background": BLUE_LIGHT, "halign": "center"}),
        ("A3:F3", {"italic": True, "font_color": "#595959", "halign": "center"}),
        ("A5:F5", {
            "bold": True, "font_color": "#FFFFFF", "background": NAVY,
            "halign": "center", "valign": "middle",
            "border": "SOLID", "border_color": BORDER,
        }),
        (f"A{total}:F{total}", {
            "bold": True, "background": BLUE_LIGHT,
            "border_top": "DOUBLE", "border_bottom": "SOLID_MEDIUM", "border_color": BORDER,
        }),
        (f"A{first}:A{last}", {"halign": "center", "border": "SOLID", "border_color": BORDER}),
        (f"B{first}:B{last}", {
            "halign": "center", "number_format": "date",
            "border": "SOLID", "border_color": BORDER,
        }),
        (f"C{first}:C{last}", {"halign": "left", "border": "SOLID", "border_color": BORDER}),
        (f"D{first}:D{last}", {"halign": "left", "wrap": True, "border": "SOLID", "border_color": BORDER}),
        (f"E{first}:E{last}", {"halign": "left", "border": "SOLID", "border_color": BORDER}),
        (f"F{first}:F{last}", {
            "halign": "right", "number_format": "money",
            "border": "SOLID", "border_color": BORDER,
        }),
    ])

    # «Зебра»: каждая вторая строка данных чуть подкрашена
    zebra_specs = [
        (f"A{r}:F{r}", {"background": ZEBRA})
        for r in range(first, last + 1) if (r - first) % 2 == 1
    ]
    if zebra_specs:
        gs.batch_format(zebra_specs)

    # Объединения: «шапка документа» и ячейка ИТОГО
    gs.merge_range("A1:F1")
    gs.merge_range("A2:F2")
    gs.merge_range("A3:F3")
    gs.merge_range(f"A{total}:E{total}")

    # Финальный «документный» вид
    gs.set_row_height(1, 32)
    gs.set_row_height(5, 26)
    gs.set_column_width("A", 45)
    gs.autofit_columns("B", "F")
    gs.freeze_rows(5)
    gs.hide_gridlines()


def write_report(
    gs: GoogleSheetsClient,
    date_from: datetime.date,
    date_to: datetime.date,
    report_type: str,
    department: str,
) -> str:
    """Создаёт новый лист и записывает в него свёрстанный отчёт.

    Аргументы:
        gs: подключённый клиент GoogleSheetsClient.
        date_from, date_to, report_type, department: параметры отчёта
            (те же, что у generate_report_data).

    Возвращает:
        Название созданного листа, напр. "Отчёт_Продажи_19.10.2026_09-30-05".
    """
    generated_at = datetime.datetime.now()
    sheet_title = (
        f"Отчёт_{report_type}_{generated_at:%d.%m.%Y_%H-%M-%S}"
    )

    data_rows = generate_report_data(date_from, date_to, report_type, department)
    matrix, layout = build_report_matrix(
        data_rows, date_from, date_to, report_type, department, generated_at,
    )

    gs.set_worksheet(sheet_title)      # создаёт лист, переключает на него
    gs.update_range(f"A1:F{layout['total_row']}", matrix)
    format_report(gs, layout)
    return sheet_title


class ReportWindow:
    """Окно Tkinter: поля отчёта, кнопка запуска, строка статуса.

    Требует установленный tkinter; на headless-сервере окно не создаётся
    (используйте generate_report_data()/write_report() напрямую).
    """

    def __init__(self, master: "tk.Tk"):
        self.master = master
        self.q: queue.Queue = queue.Queue()  # сообщения из фонового потока

        master.title("Генератор отчётов — Google Sheets")
        master.resizable(False, False)

        frame = ttk.Frame(master, padding=16)
        frame.grid(sticky="nsew")

        widgets = [
            ("Дата от (ДД.ММ.ГГГГ):", "entry"),
            ("Дата до (ДД.ММ.ГГГГ):", "entry"),
            ("Тип отчёта:", ("combo", tuple(REPORT_TYPES))),
            ("Подразделение:", ("combo", DEPARTMENTS)),
        ]
        self.entries: list = []
        for row, (label, kind) in enumerate(widgets):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            if kind == "entry":
                self.entries.append(ttk.Entry(frame, width=16))
            else:
                box = ttk.Combobox(frame, width=14, state="readonly", values=kind[1])
                box.current(0)
                self.entries.append(box)
            self.entries[-1].grid(row=row, column=1, sticky="w", pady=4)
        self.date_from_entry, self.date_to_entry, self.type_box, self.department_box = self.entries

        self.button = ttk.Button(
            frame, text="Сформировать и выгрузить в Google Sheets",
            command=self._on_generate,
        )
        self.button.grid(row=5, column=0, columnspan=2, pady=(12, 4), sticky="ew")

        # Предзаполним даты для удобства проверки
        self.date_from_entry.insert(0, "01.10.2026")
        self.date_to_entry.insert(0, "31.10.2026")

        self.status = ttk.Label(master, text="Заполните параметры и нажмите кнопку", padding=8)
        self.status.grid(sticky="ew")

        master.after(100, self._poll)

    def _poll(self) -> None:
        """Восстанавливает статус из очереди фонового потока (потокобезопасно)."""
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "status":
                    self.status.config(text=msg[1])
                elif msg[0] == "done":
                    self.status.config(text=f"Готово: создан лист «{msg[1]}»")
                    self.button.config(state="normal")
                elif msg[0] == "error":
                    messagebox.showerror("Ошибка", msg[1])
                    self.status.config(text="Ошибка (см. окно сообщения)")
                    self.button.config(state="normal")
        except queue.Empty:
            pass
        self.master.after(100, self._poll)

    def _on_generate(self) -> None:
        """Проверяет поля и запускает выгрузку в фоновом потоке."""
        try:
            date_from = parse_date(self.date_from_entry.get())
            date_to = parse_date(self.date_to_entry.get())
            report_type = self.type_box.get()
            department = self.department_box.get()
            if date_to < date_from:
                raise ValueError("Дата «до» не может быть раньше даты «от»")
        except ValueError as err:
            messagebox.showwarning("Проверка формы", str(err))
            return

        self.button.config(state="disabled")
        self.q.put(("status", "Формирую отчёт и записываю в Google Sheets…"))
        threading.Thread(target=self._worker, args=(date_from, date_to, report_type, department), daemon=True).start()

    def _worker(self, date_from, date_to, report_type, department) -> None:
        """Фоновая выгрузка: сеть не блокирует интерфейс."""
        try:
            gs = GoogleSheetsClient()
            sheet_title = write_report(gs, date_from, date_to, report_type, department)
        except Exception as err:  # ошибки сети/API показываются в окне
            self.q.put(("error", f"Не удалось записать отчёт:\n{err}"))
        else:
            self.q.put(("done", sheet_title))


def main() -> None:
    """Точка входа: создаёт окно и запускает обработку событий Tk."""
    if tk is None:
        raise RuntimeError(
            "tkinter недоступен. На локальном компьютере установите Python "
            "с tkinter (windows-инсталлятор содержит его по умолчанию; "
            "в Debian/Ubuntu: sudo apt install python3-tk)."
        )
    root = tk.Tk()
    ReportWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
