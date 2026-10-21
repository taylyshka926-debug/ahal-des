import json
import calendar
import os
from datetime import datetime

from kivy.app import App
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView

Window.clearcolor = (0.07, 0.08, 0.12, 1)

MONTHS = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

# Norms transcribed from the user's 2026 Excel workbook.
DEFAULT_NORMS = [168, 160, 151, 176, 135, 176, 184, 168, 160, 167, 168, 182]

# Imported from the user's 2026 Excel shift calendar. D=day shift (11 h), N=night shift (4 day + 7 night), -=off.
# Exact day-by-day shift calendar read from the user Excel workbook (portative sheet).
# D = 11-hour day shift, N = 4/7 shift (4 day-rate hours + 7 night-rate hours), - = off.
SCHEDULE_2026 = {
    '1': ['DD--NN--DD--NN--DD--NN--DD--NN-', 'NN--DD--NN--DD--NN--DD--NN--DD-', '--DD--NN--DD--NN--DD--NN--DD--N', '--NN--DD--NN--DD--NN--DD--NN--D'],
    '2': ['-DD--NN--DD--NN--DD--NN--DD-', '-NN--DD--NN--DD--NN--DD--NN-', 'N--DD--NN--DD--NN--DD--NN--D', 'D--NN--DD--NN--DD--NN--DD--N'],
    '3': ['-NN--DD--NN--DD--NN--DD--NN--DD', '-DD--NN--DD--NN--DD--NN--DD--NN', 'D--NN--DD--NN--DD--NN--DD--NN--', 'N--DD--NN--DD--NN--DD--NN--DD--'],
    '4': ['--NN--DD--NN--DD--NN--DD--NN--', '--DD--NN--DD--NN--DD--NN--DD--', 'DD--NN--DD--NN--DD--NN--DD--NN', 'NN--DD--NN--DD--NN--DD--NN--DD'],
    '5': ['DD--NN--DD--NN--DD--NN--DD--NN-', 'NN--DD--NN--DD--NN--DD--NN--DD-', '--DD--NN--DD--NN--DD--NN--DD--N', '--NN--DD--NN--DD--NN--DD--NN--D'],
    '6': ['-DD--NN--DD--NN--DD--NN--DD--N', '-NN--DD--NN--DD--NN--DD--NN--D', 'N--DD--NN--DD--NN--DD--NN--DD-', 'D--NN--DD--NN--DD--NN--DD--NN-'],
    '7': ['N--DD--NN--DD--NN--DD--NN--DD--', 'D--NN--DD--NN--DD--NN--DD--NN--', '-NN--DD--NN--DD--NN--DD--NN--DD', '-DD--NN--DD--NN--DD--NN--DD--NN'],
    '8': ['NN--DD--NN--DD--NN--DD--NN--DD-', 'DD--NN--DD--NN--DD--NN--DD--NN-', '--NN--DD--NN--DD--NN--DD--NN--D', '--DD--NN--DD--NN--DD--NN--DD--N'],
    '9': ['-NN--DD--NN--DD--NN--DD--NN--D', '-DD--NN--DD--NN--DD--NN--DD--N', 'D--NN--DD--NN--DD--NN--DD--NN-', 'N--DD--NN--DD--NN--DD--NN--DD-'],
    '10': ['D--NN--DD--NN--DD--NN--DD--NN--', 'N--DD--NN--DD--NN--DD--NN--DD--', '-DD--NN--DD--NN--DD--NN--DD--NN', '-NN--DD--NN--DD--NN--DD--NN--DD'],
    '11': ['DD--NN--DD--NN--DD--NN--DD--NN', 'NN--DD--NN--DD--NN--DD--NN--DD', '--DD--NN--DD--NN--DD--NN--DD--', '--NN--DD--NN--DD--NN--DD--NN--'],
    '12': ['--DD--NN--DD--NN--DD--NN--DD--N', '--NN--DD--NN--DD--NN--DD--NN--D', 'NN--DD--NN--DD--NN--DD--NN--DD-', 'DD--NN--DD--NN--DD--NN--DD--NN-'],
}

# Public-holiday dates used by the user's 2026 Excel formulas.
# Exact holiday-hour totals are read from cached results in the portative sheet.
HOLIDAY_DAYS_2026 = {
    1: {1},
    3: {8, 20, 21, 22},
    5: {18, 27, 28, 29},
    9: {27},
    10: {6},
    12: {12},
}

HOLIDAY_HOURS_2026 = {
    1:  [11, 4, 0, 0],
    2:  [0, 0, 0, 0],
    3:  [18, 11, 33, 26],
    4:  [0, 0, 0, 0],
    5:  [15, 29, 22, 22],
    6:  [0, 0, 0, 0],
    7:  [0, 0, 0, 0],
    8:  [0, 0, 0, 0],
    9:  [11, 11, 0, 0],
    10: [7, 0, 4, 11],
    11: [0, 0, 0, 0],
    12: [11, 11, 0, 0],
}

# Graph selections shown in the user's 2026 payroll sheet. For months without
# an explicit selection, Graph 1 is the initial choice and remains editable.
EXCEL_DEFAULT_GRAPH_2026 = {
    1: 4, 2: 4, 3: 4, 4: 1, 5: 1, 6: 1,
    7: 1, 8: 2, 9: 3, 10: 3, 11: 1, 12: 1,
}

WHITE = (0.95, 0.96, 0.99, 1)
MUTED = (0.72, 0.76, 0.84, 1)


class AhalDESApp(App):
    title = "Ahal DES"

    def build(self):
        self.filename = os.path.join(self.user_data_dir, "ahal_des_data.json")
        self.data = self.load_data()

        root = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        header = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(72), spacing=dp(2))
        header.add_widget(Label(
            text="[b]Ahal DES[/b]", markup=True, font_size="28sp",
            color=WHITE, size_hint_y=None, height=dp(40)
        ))
        header.add_widget(Label(
            text="РАСЧЁТ ЗАРАБОТНОЙ ПЛАТЫ", font_size="13sp",
            color=MUTED, size_hint_y=None, height=dp(24)
        ))
        root.add_widget(header)

        scroll = ScrollView()
        content = BoxLayout(orientation="vertical", spacing=dp(8), size_hint_y=None)
        content.bind(minimum_height=content.setter("height"))

        content.add_widget(self.section_label("Период расчёта"))
        row = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(8))
        self.year = TextInput(
            text="2026", hint_text="Год", multiline=False,
            input_filter="int", input_type="number", size_hint_x=0.35
        )
        self.month = Spinner(
            text=MONTHS[0], values=MONTHS, size_hint_x=0.65
        )
        self.month.bind(text=self.on_month_change)
        self.year.bind(text=self.on_year_or_month_change)
        row.add_widget(self.year)
        row.add_widget(self.month)
        content.add_widget(row)

        content.add_widget(self.section_label("Оклад и часы"))
        self.salary = self.add_field(content, "Месячный оклад по Excel (манаты)", "3421")
        self.norm = self.add_field(content, "Норма часов выбранного месяца", str(DEFAULT_NORMS[0]))
        self.night = self.add_field(content, "Ночные часы (из Excel)", "0")
        self.holiday = self.add_field(content, "Праздничные часы (из Excel)", "0")

        content.add_widget(self.section_label("Нормы часов на весь год"))
        content.add_widget(Label(
            text="Можно исправить любую норму. Для расчёта берётся норма выбранного месяца.",
            color=MUTED, font_size="12sp", size_hint_y=None, height=dp(38),
            halign="left", valign="middle"
        ))
        self.norm_fields = []
        norms_grid = GridLayout(cols=2, spacing=dp(6), size_hint_y=None)
        norms_grid.bind(minimum_height=norms_grid.setter("height"))
        for i, month_name in enumerate(MONTHS):
            cell = BoxLayout(orientation="vertical", spacing=dp(2),
                             size_hint_y=None, height=dp(58))
            cell.add_widget(Label(
                text=month_name, color=WHITE, font_size="13sp",
                size_hint_y=None, height=dp(22), halign="left"
            ))
            field = TextInput(
                text=str(DEFAULT_NORMS[i]), multiline=False,
                input_filter="int", input_type="number",
                size_hint_y=None, height=dp(34)
            )
            field.bind(text=lambda inst, value, idx=i: self.on_norm_field_change(idx, value))
            self.norm_fields.append(field)
            cell.add_widget(field)
            norms_grid.add_widget(cell)
        content.add_widget(norms_grid)

        content.add_widget(self.section_label("Автоматический календарь смен 2026"))
        # Separate calendar selector makes it clear that all 12 months are stored,
        # while only the currently selected month is displayed.
        calendar_month_row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(6))
        calendar_month_row.add_widget(Button(text="‹", size_hint_x=0.16, on_press=self.previous_calendar_month))
        self.calendar_month = Spinner(text=MONTHS[0], values=MONTHS, size_hint_x=0.68)
        self.calendar_month.bind(text=self.on_calendar_month_change)
        calendar_month_row.add_widget(self.calendar_month)
        calendar_month_row.add_widget(Button(text="›", size_hint_x=0.16, on_press=self.next_calendar_month))
        content.add_widget(calendar_month_row)

        schedule_row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(8))
        schedule_row.add_widget(Label(text="График из Excel:", color=WHITE, font_size="13sp", size_hint_x=0.55))
        self.schedule_team = Spinner(text="График 1", values=("График 1", "График 2", "График 3", "График 4"), size_hint_x=0.45)
        self.schedule_team.bind(text=self.refresh_calendar)
        schedule_row.add_widget(self.schedule_team)
        content.add_widget(schedule_row)
        self.calendar_summary = Label(text="Выбери месяц и график.", color=MUTED, font_size="12sp", size_hint_y=None, height=dp(42), halign="left", valign="middle")
        self.calendar_summary.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], None)))
        content.add_widget(self.calendar_summary)
        # Календарь показывает только выбранный месяц; переключение выполняется через список месяца выше.
        self.calendar_grid = BoxLayout(orientation="vertical", spacing=dp(10), size_hint_y=None)
        self.calendar_grid.bind(minimum_height=self.calendar_grid.setter("height"))
        content.add_widget(self.calendar_grid)
        self.apply_schedule_button = Button(text="Подставить смены и праздники", size_hint_y=None, height=dp(44), on_press=self.apply_schedule)
        content.add_widget(self.apply_schedule_button)

        note = Label(
            text="Календарь и праздничные часы взяты из твоего Excel за 2026 год. Д — дневная смена 11 ч; Н — смена 4 дневных + 7 ночных; ★ — праздничная дата.",
            color=MUTED, font_size="12sp", size_hint_y=None, height=dp(66),
            halign="left", valign="middle"
        )
        note.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], None)))
        content.add_widget(note)

        buttons = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=dp(76))
        buttons.add_widget(Button(text="Расчёт за месяц", on_press=self.calculate))
        buttons.add_widget(Button(text="Итог за год", on_press=self.calculate_year))
        buttons.add_widget(Button(text="Сохранить месяц", on_press=self.save_month))
        buttons.add_widget(Button(text="Загрузить месяц", on_press=self.load_month))
        buttons.add_widget(Button(text="Сохранить нормы года", on_press=self.save_year_norms))
        buttons.add_widget(Button(text="Итог за год", on_press=self.calculate_year))
        content.add_widget(buttons)

        content.add_widget(self.section_label("Результат расчёта"))
        self.result = Label(
            text="Введи ночные и праздничные часы, затем нажми «Рассчитать».",
            color=WHITE, font_size="14sp", size_hint_y=None, height=dp(300),
            halign="left", valign="top"
        )
        self.result.bind(size=lambda inst, val: setattr(inst, "text_size", (val[0], None)))
        content.add_widget(self.result)

        scroll.add_widget(content)
        root.add_widget(scroll)

        self.load_year_norms()
        self.on_month_change(self.month, self.month.text)
        self.calendar_month.text = self.month.text
        self.load_month(None)
        self.refresh_calendar()
        return root

    def selected_month_index(self):
        # The calendar selector is authoritative for the displayed schedule.
        month_text = self.calendar_month.text if hasattr(self, "calendar_month") else self.month.text
        return MONTHS.index(month_text) + 1

    def on_calendar_month_change(self, spinner, text):
        # Keep the salary period and calendar month synchronized.
        if hasattr(self, "month") and self.month.text != text:
            self.month.text = text
        if hasattr(self, "norm_fields") and hasattr(self, "norm"):
            idx = MONTHS.index(text)
            self.norm.text = self.norm_fields[idx].text or str(DEFAULT_NORMS[idx])
        if hasattr(self, "schedule_team"):
            month_idx = MONTHS.index(text) + 1
            record = self.data.get(f"{self.year.text.strip() or '2026'}-{text}", {})
            graph = int(record.get("graph", EXCEL_DEFAULT_GRAPH_2026.get(month_idx, 1)))
            if self.schedule_team.text != f"График {graph}":
                self.schedule_team.text = f"График {graph}"
        if hasattr(self, "calendar_grid"):
            self.refresh_calendar()

    def previous_calendar_month(self, instance):
        idx = (MONTHS.index(self.calendar_month.text) - 1) % len(MONTHS)
        self.calendar_month.text = MONTHS[idx]

    def next_calendar_month(self, instance):
        idx = (MONTHS.index(self.calendar_month.text) + 1) % len(MONTHS)
        self.calendar_month.text = MONTHS[idx]

    def schedule_codes(self):
        month_idx = self.selected_month_index()
        year = self.year.text.strip() or "2026"
        if year != "2026":
            return None
        team_idx = int(self.schedule_team.text.split()[-1]) - 1
        return SCHEDULE_2026[str(month_idx)][team_idx]

    def refresh_calendar(self, *args):
        if not hasattr(self, "calendar_grid"):
            return
        self.calendar_grid.clear_widgets()
        try:
            year = int(self.year.text or "2026")
            month_idx = self.selected_month_index()
            month_name = MONTHS[month_idx - 1]
            if year != 2026:
                self.calendar_summary.text = "График из Excel есть только за 2026 год. Для другого года календарь нужно заполнить отдельно."
                return

            team_idx = int(self.schedule_team.text.split()[-1]) - 1
            codes = SCHEDULE_2026[str(month_idx)][team_idx]
            first_weekday, days_count = calendar.monthrange(year, month_idx)
            month_grid = GridLayout(cols=7, spacing=dp(3), size_hint_y=None)
            month_grid.bind(minimum_height=month_grid.setter("height"))
            for weekday in ("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"):
                month_grid.add_widget(Label(text=weekday, color=MUTED, font_size="11sp", size_hint_y=None, height=dp(22)))
            for _ in range(first_weekday):
                month_grid.add_widget(Label(text="", size_hint_y=None, height=dp(48)))

            day_count = night_count = 0
            holiday_days = HOLIDAY_DAYS_2026.get(month_idx, set())
            month_days = min(days_count, len(codes))
            for day in range(1, month_days + 1):
                code = codes[day - 1]
                is_holiday = day in holiday_days
                if code == "D":
                    label = f"{day}\nД" + (" ★" if is_holiday else "")
                    day_count += 1
                    color = (0.72, 0.12, 0.12, 1)  # red: work shift
                elif code == "N":
                    label = f"{day}\nН" + (" ★" if is_holiday else "")
                    night_count += 1
                    color = (0.72, 0.12, 0.12, 1)  # red: work shift
                else:
                    label = f"{day}\n" + ("ПР ★" if is_holiday else "ОТД")
                    color = (0.92, 0.76, 0.20, 1)  # yellow: day off
                if is_holiday and code != "-":
                    color = (0.84, 0.25, 0.12, 1)
                month_grid.add_widget(Button(
                    text=label, font_size="12sp", size_hint_y=None, height=dp(48),
                    background_normal="", background_color=color, disabled=True
                ))
            for _ in range((7 - ((first_weekday + month_days) % 7)) % 7):
                month_grid.add_widget(Label(text="", size_hint_y=None, height=dp(48)))

            self.calendar_grid.add_widget(Label(
                text=f"[b]{month_name} {year}[/b]", markup=True, color=WHITE,
                font_size="18sp", size_hint_y=None, height=dp(32), halign="left", valign="middle"
            ))
            self.calendar_grid.add_widget(month_grid)
            graph_idx = int(self.schedule_team.text.split()[-1]) - 1
            holiday_hours = HOLIDAY_HOURS_2026.get(month_idx, [0, 0, 0, 0])[graph_idx]
            holiday_dates_text = ", ".join(str(d) for d in sorted(holiday_days)) or "нет"
            self.calendar_summary.text = (
                f"{month_name} {year} · график {self.schedule_team.text[-1]}. "
                f"Дневных смен: {day_count}, ночных смен: {night_count}. "
                f"Праздничные даты: {holiday_dates_text}. Часы по Excel: {holiday_hours}."
            )
        except (ValueError, KeyError, IndexError):
            self.calendar_summary.text = "Не удалось показать календарь. Проверь год и график."

    def apply_schedule(self, instance):
        codes = self.schedule_codes()
        if codes is None:
            self.calendar_summary.text = "Автоматический график из Excel доступен только для 2026 года."
            return
        month_idx = self.selected_month_index()
        graph_idx = int(self.schedule_team.text.split()[-1]) - 1
        day_count = codes.count("D")
        night_count = codes.count("N")
        # Each 4/7 shift contributes 4 day-rate hours and 7 night hours.
        self.night.text = str(night_count * 7)
        holiday_hours = HOLIDAY_HOURS_2026.get(month_idx, [0, 0, 0, 0])[graph_idx]
        self.holiday.text = str(holiday_hours)
        total_shifts = day_count + night_count
        total_hours = total_shifts * 11
        day_rate_hours = day_count * 11 + night_count * 4
        holiday_dates_text = ", ".join(str(d) for d in sorted(HOLIDAY_DAYS_2026.get(month_idx, set()))) or "нет"
        self.calendar_summary.text = (
            f"Подставлено из Excel: {day_count} дневных и {night_count} ночных смен. "
            f"Всего {total_shifts} смен / {total_hours} ч. "
            f"Дневные часы: {day_rate_hours}; ночные: {night_count * 7}. "
            f"Праздничные даты: {holiday_dates_text}; праздничные часы: {holiday_hours}."
        )
        # Keep generated night hours in the current month's saved record too.
        self.data.setdefault(self.key(), {})["night"] = self.night.text
        self.data[self.key()]["salary"] = self.salary.text or "3421"
        self.data[self.key()]["holiday"] = self.holiday.text or "0"
        self.data[self.key()]["graph"] = graph_idx + 1
        self.save_data_file()

    def section_label(self, text):
        return Label(
            text=f"[b]{text}[/b]", markup=True, color=WHITE, font_size="16sp",
            size_hint_y=None, height=dp(30), halign="left", valign="middle"
        )

    def add_field(self, parent, title, default):
        parent.add_widget(Label(
            text=title, color=WHITE, font_size="13sp",
            size_hint_y=None, height=dp(26), halign="left", valign="middle"
        ))
        field = TextInput(
            text=default, multiline=False, input_filter="float",
            input_type="number", size_hint_y=None, height=dp(44)
        )
        parent.add_widget(field)
        return field

    def load_data(self):
        try:
            with open(self.filename, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def save_data_file(self):
        with open(self.filename, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def key(self):
        year = self.year.text.strip() or "2026"
        return f"{year}-{self.month.text}"

    def year_key(self):
        return self.year.text.strip() or "2026"

    def on_year_or_month_change(self, *args):
        if hasattr(self, "norm_fields"):
            self.load_year_norms()
            self.on_month_change(self.month, self.month.text)
        if hasattr(self, "calendar_month") and self.calendar_month.text != self.month.text:
            self.calendar_month.text = self.month.text
        if hasattr(self, "calendar_grid"):
            self.refresh_calendar()

    def on_month_change(self, spinner, text):
        if not hasattr(self, "norm"):
            return
        idx = MONTHS.index(text)
        if self.norm_fields:
            self.norm.text = self.norm_fields[idx].text
        record = self.data.get(self.key(), {})
        self.night.text = str(record.get("night", "0"))
        self.holiday.text = str(record.get("holiday", "0"))
        if record:
            self.salary.text = str(record.get("salary", "3421"))
        if hasattr(self, "calendar_month") and self.calendar_month.text != text:
            self.calendar_month.text = text
        if hasattr(self, "calendar_grid"):
            self.refresh_calendar()

    def on_norm_field_change(self, idx, value):
        if hasattr(self, "month") and MONTHS.index(self.month.text) == idx and hasattr(self, "norm"):
            if value.strip():
                self.norm.text = value

    def load_year_norms(self):
        if not hasattr(self, "norm_fields"):
            return
        year_record = self.data.get("norms-" + self.year_key(), {})
        values = year_record.get("values", DEFAULT_NORMS)
        for i, field in enumerate(self.norm_fields):
            if not field.focus:
                field.text = str(values[i] if i < len(values) else DEFAULT_NORMS[i])

    def numbers(self):
        salary = float(self.salary.text.replace(",", ".") or 0)
        norm = float(self.norm.text.replace(",", ".") or 0)
        night_hours = float(self.night.text.replace(",", ".") or 0)
        holiday_hours = float(self.holiday.text.replace(",", ".") or 0)
        if salary < 0 or norm <= 0 or night_hours < 0 or holiday_hours < 0:
            raise ValueError("Проверь оклад и часы")
        return salary, norm, night_hours, holiday_hours

    def month_shift_info(self, month_idx, graph_num=None):
        """Return shift counts/hours for a month from the Excel-derived 2026 roster."""
        if graph_num is None:
            graph_num = EXCEL_DEFAULT_GRAPH_2026.get(month_idx, 1)
        codes_list = SCHEDULE_2026.get(str(month_idx))
        if not codes_list or not 1 <= int(graph_num) <= 4:
            return None
        codes = codes_list[int(graph_num) - 1]
        day_shifts = codes.count("D")
        night_shifts = codes.count("N")
        return {
            "codes": codes,
            "graph": int(graph_num),
            "day_shifts": day_shifts,
            "night_shifts": night_shifts,
            "total_hours": (day_shifts + night_shifts) * 11,
            "day_hours": day_shifts * 11 + night_shifts * 4,
            "night_hours": night_shifts * 7,
        }

    def hours_balance_before(self, month_idx, current_graph=None, current_holiday=None):
        """Excel's cumulative rest-hour balance starts at 42 hours before January."""
        balance = 42.0
        for idx in range(1, month_idx + 1):
            record = self.data.get(f"{self.year.text}-{MONTHS[idx - 1]}", {})
            graph = (current_graph if idx == month_idx and current_graph is not None
                     else int(record.get("graph", EXCEL_DEFAULT_GRAPH_2026.get(idx, 1))))
            info = self.month_shift_info(idx, graph)
            if info is None:
                continue
            norm = float(record.get("norm", self.norm_fields[idx - 1].text or DEFAULT_NORMS[idx - 1]))
            holiday = (current_holiday if idx == month_idx and current_holiday is not None
                       else float(record.get("holiday", HOLIDAY_HOURS_2026.get(idx, [0, 0, 0, 0])[graph - 1])))
            balance += info["total_hours"] - norm - holiday
        return balance

    def calculate(self, instance):
        try:
            salary, norm, night_hours, holiday_hours = self.numbers()
            hourly_rate = salary / norm

            # Shift counts and hours come from the selected Excel roster.
            month_idx = self.selected_month_index()
            graph_num = int(self.schedule_team.text.split()[-1])
            info = self.month_shift_info(month_idx, graph_num) if (self.year.text.strip() or "2026") == "2026" else None
            if info is not None:
                day_shifts = info["day_shifts"]
                night_shifts = info["night_shifts"]
                day_hours = info["day_hours"]
                scheduled_night_hours = info["night_hours"]
                total_hours = info["total_hours"]
                deviation = total_hours - norm - holiday_hours
                balance = self.hours_balance_before(month_idx, graph_num, holiday_hours)
                balance_before = balance - deviation
                shift_lines = (
                    f"График: {graph_num}-я вахта\n"
                    f"Дневные смены: {day_shifts} × 11 ч = {day_shifts * 11} ч\n"
                    f"Ночные смены: {night_shifts} × 11 ч = {night_shifts * 11} ч\n"
                    f"  из них дневная ставка: {night_shifts} × 4 ч = {night_shifts * 4} ч\n"
                    f"  ночная ставка: {night_shifts} × 7 ч = {scheduled_night_hours} ч\n"
                    f"Дневные часы всего: {day_hours} ч; ночные часы: {scheduled_night_hours} ч\n"
                    f"Всего часов по вахтам: {total_hours} ч\n"
                    f"Разница с нормой: {total_hours} − {norm:g} − праздничные {holiday_hours:g} = {deviation:+g} ч\n"
                    f"Остаток отгула: {balance_before:g} + ({deviation:+g}) = {balance:g} ч\n"
                )
                if abs(night_hours - scheduled_night_hours) > 0.001:
                    shift_lines += f"ВНИМАНИЕ: введено ночных часов {night_hours:g}, а по графику {scheduled_night_hours} ч.\n"
            else:
                shift_lines = "Смены, переработка и остаток отгула по календарю доступны только за 2026 год.\n"

            night_pay = hourly_rate * night_hours * 0.50
            holiday_pay = hourly_rate * holiday_hours * 2.0
            gross = salary + night_pay + holiday_pay

            pension = gross * 0.02
            income_tax = (gross - 50 - pension) * 0.10
            union = gross * 0.01
            youth = gross * 0.01
            alimony = (gross - income_tax) * 0.25
            fixed = 5.0
            deductions = income_tax + pension + union + youth + alimony + fixed
            net = gross - deductions

            lines = [
                f"СПИСОК РАСЧЁТА: {self.month.text} {self.year.text}",
                "=" * 34,
                f"1. Месячный оклад: {salary:.2f} ман.",
                f"2. Норма часов: {norm:g} ч.",
                f"3. Часовая ставка: {salary:.2f} ÷ {norm:g} = {hourly_rate:.4f} ман.",
                "",
                "СМЕНЫ И РАБОЧИЕ ЧАСЫ",
                shift_lines.rstrip(),
                "",
                "НАЧИСЛЕНИЯ",
                f"4. Ночные часы: {night_hours:g} × {hourly_rate:.4f} × 50% = +{night_pay:.2f} ман.",
                f"5. Праздничные часы: {holiday_hours:g} × {hourly_rate:.4f} × 200% = +{holiday_pay:.2f} ман.",
                f"6. Начислено всего: {salary:.2f} + {night_pay:.2f} + {holiday_pay:.2f} = {gross:.2f} ман.",
                "",
                "УДЕРЖАНИЯ",
                f"7. Пенсионный (2%): −{pension:.2f} ман.",
                f"8. Подоходный налог (10%): −{income_tax:.2f} ман.",
                f"9. Профсоюзный (1%): −{union:.2f} ман.",
                f"10. Молодёжный (1%): −{youth:.2f} ман.",
                f"11. Алименты (25% от суммы после налога): −{alimony:.2f} ман.",
                f"12. Abatlandyryş: −{fixed:.2f} ман.",
                f"Всего удержаний: {deductions:.2f} ман.",
                "",
                f"К ВЫПЛАТЕ: {net:.2f} ман."
            ]
            self.result.markup = False
            self.result.text = "\n".join(lines)
            self.result.height = dp(max(300, 38 * len(lines)))
        except (ValueError, ZeroDivisionError, TypeError):
            self.result.markup = False
            self.result.text = "Проверь числа. Норма часов должна быть больше нуля."

    def calculate_year(self, instance):
        """Показывает список расчётов за все 12 месяцев и годовые итоги."""
        try:
            # Сначала берём введённые сейчас значения выбранного месяца.
            current_salary, current_norm, current_night, current_holiday = self.numbers()
            current_idx = MONTHS.index(self.month.text)
            norms = []
            for i, field in enumerate(self.norm_fields):
                value = float((field.text or str(DEFAULT_NORMS[i])).replace(",", "."))
                if value <= 0:
                    raise ValueError("Норма часов должна быть больше нуля")
                norms.append(value)
            norms[current_idx] = current_norm

            lines = [f"СПИСОК РАСЧЁТА ЗА {self.year.text}", "=" * 29]
            totals = {"gross": 0.0, "deductions": 0.0, "net": 0.0}
            for i, month_name in enumerate(MONTHS):
                record = self.data.get(f"{self.year.text}-{month_name}", {})
                if i == current_idx:
                    salary, night, holiday = current_salary, current_night, current_holiday
                    graph = int(self.schedule_team.text.split()[-1])
                else:
                    salary = float(record.get("salary", current_salary))
                    graph = int(record.get("graph", EXCEL_DEFAULT_GRAPH_2026.get(i + 1, 1)))
                    default_holiday = HOLIDAY_HOURS_2026.get(i + 1, [0, 0, 0, 0])[graph - 1]
                    holiday = float(record.get("holiday", default_holiday))
                    info_saved = self.month_shift_info(i + 1, graph) if self.year.text.strip() == "2026" else None
                    night = float(record.get("night", info_saved["night_hours"] if info_saved else 0))
                norm = norms[i]
                info = self.month_shift_info(i + 1, graph) if self.year.text.strip() == "2026" else None
                if info:
                    actual_hours = info["total_hours"]
                    deviation = actual_hours - norm - holiday
                    balance_before = totals.get("hours_balance", 42.0) if i == 0 else totals["hours_balance"]
                    totals["hours_balance"] = balance_before + deviation
                else:
                    actual_hours = None
                    deviation = None
                    balance_before = None
                rate = salary / norm
                night_pay = rate * night * 0.50
                holiday_pay = rate * holiday * 2.0
                gross = salary + night_pay + holiday_pay
                pension = gross * 0.02
                tax = (gross - 50 - pension) * 0.10
                union = gross * 0.01
                youth = gross * 0.01
                alimony = (gross - tax) * 0.25
                fixed = 5.0
                deductions = tax + pension + union + youth + alimony + fixed
                net = gross - deductions
                totals["gross"] += gross
                totals["deductions"] += deductions
                totals["net"] += net
                lines.append(
                    f"\n{month_name}:\n"
                    f"  Оклад: {salary:.2f} ман. | Норма: {norm:g} ч.\n"
                    f"  Ночные: {night:g} ч. (+{night_pay:.2f})\n"
                    f"  Праздничные: {holiday:g} ч. (+{holiday_pay:.2f})\n"
                    + (f"  Вахта: график {graph}; фактически {actual_hours:g} ч.; разница {deviation:+g} ч.; остаток отгула {totals['hours_balance']:g} ч.\n" if info else "  График/остаток отгула: только для 2026 года.\n")
                    + f"  Начислено: {gross:.2f} ман.\n"
                    f"  Удержания: {deductions:.2f} ман.\n"
                    f"  К выплате: {net:.2f} ман."
                )
            lines.extend([
                "\n" + "=" * 29,
                f"ИТОГО НАЧИСЛЕНО: {totals['gross']:.2f} ман.",
                f"ИТОГО УДЕРЖАНИЙ: {totals['deductions']:.2f} ман.",
                f"ИТОГО ЗА ГОД К ВЫПЛАТЕ: {totals['net']:.2f} ман.",
                "\nОстаток отгула на начало года по Excel: 42 ч. Разница за месяц = часы вахт − норма − праздничные часы.\nВажно: проверь график и праздничные часы каждого месяца перед использованием для зарплаты."
            ])
            self.result.markup = False
            self.result.text = "\n".join(lines)
            self.result.height = dp(max(300, 110 + len(lines) * 27))
        except (ValueError, ZeroDivisionError, TypeError):
            self.result.markup = False
            self.result.text = "Проверь оклад, нормы и часы. Норма каждого месяца должна быть больше нуля."

    def save_month(self, instance):
        try:
            salary, norm, night, holiday = self.numbers()
            self.data[self.key()] = {
                "salary": salary, "norm": norm, "night": night, "holiday": holiday,
                "graph": int(self.schedule_team.text.split()[-1])
            }
            self.save_year_norms_to_data()
            self.save_data_file()
            self.result.text = f"Данные за {self.month.text} {self.year.text} сохранены на телефоне."
        except (ValueError, OSError):
            self.result.text = "Не удалось сохранить. Проверь введённые значения."

    def load_month(self, instance):
        record = self.data.get(self.key(), {})
        self.salary.text = str(record.get("salary", 3421))
        idx = MONTHS.index(self.month.text)
        self.norm.text = str(record.get("norm", self.norm_fields[idx].text or DEFAULT_NORMS[idx]))
        self.night.text = str(record.get("night", 0))
        self.holiday.text = str(record.get("holiday", 0))
        if hasattr(self, "schedule_team"):
            graph = int(record.get("graph", EXCEL_DEFAULT_GRAPH_2026.get(idx + 1, 1)))
            self.schedule_team.text = f"График {graph}"
        if hasattr(self, "result"):
            self.result.markup = False
            self.result.text = f"Загружены данные: {self.month.text} {self.year.text}."

    def save_year_norms_to_data(self):
        values = []
        for i, field in enumerate(self.norm_fields):
            value = int(field.text or DEFAULT_NORMS[i])
            if value <= 0:
                raise ValueError("Норма должна быть больше нуля")
            values.append(value)
        self.data["norms-" + self.year_key()] = {"values": values}

    def save_year_norms(self, instance):
        try:
            self.save_year_norms_to_data()
            self.save_data_file()
            self.result.markup = False
            self.result.text = f"Все 12 норм часов за {self.year_key()} сохранены на телефоне."
        except (ValueError, OSError):
            self.result.markup = False
            self.result.text = "Не удалось сохранить нормы. Проверь, чтобы все нормы были больше нуля."


if __name__ == "__main__":
    AhalDESApp().run()
