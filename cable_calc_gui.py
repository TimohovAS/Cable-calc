import json
import logging
import math
import typing
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape
from zipfile import ZipFile
import os

from openpyxl import Workbook
from openpyxl.utils import get_column_letter


class Tooltip:
    def __init__(self, widget: tk.Widget, text_getter: typing.Callable[[], str]) -> None:
        self.widget = widget
        self.text_getter = text_getter
        self.tipwindow: tk.Toplevel | None = None
        self.widget.bind("<Enter>", self._show_tip)
        self.widget.bind("<Leave>", self._hide_tip)

    def _show_tip(self, event: tk.Event | None) -> None:
        text = self.text_getter()
        if not text:
            return
        if self.tipwindow is not None:
            self._hide_tip(None)
        x = y = 0
        if event is not None:
            x = event.x_root + 12
            y = event.y_root + 8
        else:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + 8
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=text,
            justify=tk.LEFT,
            background="#ffffe0",
            relief=tk.SOLID,
            borderwidth=1,
            font=("TkDefaultFont", 9),
            wraplength=280,
        )
        label.pack(ipadx=4, ipady=2)

    def _hide_tip(self, _: tk.Event | None) -> None:
        if self.tipwindow is not None:
            self.tipwindow.destroy()
            self.tipwindow = None


logging.basicConfig(
    filename="cable_calc.log",
    level=logging.ERROR,
    format="%(asctime)s %(levelname)s %(message)s",
)


class CableCalcApp(tk.Tk):
    WINDOW_TITLE_KEY = "app.title"
    WINDOW_GEOMETRY = "1200x800"
    DEFAULT_LANGUAGE = "ru"


    LANGUAGES = {"ru": "Русский", "sr": "Srpski", "en": "English"}
    DEFAULT_MEDIUM = "air"

    TRANSLATIONS = {
        "app.title": {"ru": "Расчёт кабелей", "sr": "Proračun kablova", "en": "Cable calculation"},
        "menu.file": {"ru": "Файл", "sr": "Datoteka", "en": "File"},
        "menu.language": {"ru": "Язык", "sr": "Jezik", "en": "Language"},
        "menu.save_project": {"ru": "Сохранить проект…", "sr": "Sačuvaj projekat…", "en": "Save project…"},
        "menu.load_project": {"ru": "Загрузить проект…", "sr": "Učitaj projekat…", "en": "Load project…"},
        "menu.export_excel": {"ru": "Экспорт в Excel…", "sr": "Izvoz u Excel…", "en": "Export to Excel…"},
        "tab.calculation": {"ru": "Расчёт", "sr": "Proračun", "en": "Calculation"},
        "tab.help": {"ru": "Помощь", "sr": "Pomoć", "en": "Help"},
        "frame.input": {"ru": "Ввод данных", "sr": "Unos podataka", "en": "Input"},
        "frame.intermediate": {"ru": "Промежуточные результаты", "sr": "Međurezultati", "en": "Intermediate results"},
        "frame.table": {"ru": "Результаты расчёта", "sr": "Rezultati proračuna", "en": "Calculation results"},
        "button.add_row": {
            "ru": "Рассчитать и добавить строку",
            "sr": "Izračunaj i dodaj red",
            "en": "Calculate and add row",
        },
        "button.clear": {"ru": "Очистить список", "sr": "Očisti listu", "en": "Clear list"},
        "button.select_optimal": {
            "ru": "Подобрать параметры",
            "sr": "Odaberi parametre",
            "en": "Select parameters",
        },
        "button.remove_row": {
            "ru": "Удалить выбранную строку",
            "sr": "Obriši izabrani red",
            "en": "Remove selected row",
        },
        "button.load_row": {
            "ru": "Загрузить строку в форму",
            "sr": "Učitaj red u formu",
            "en": "Load row into form",
        },
        "label.circuit": {"ru": "Цепь", "sr": "Strujni krug", "en": "Circuit"},
        "label.segment_from": {"ru": "Участок ОТ", "sr": "Deonica OD", "en": "Segment from"},
        "label.segment_to": {"ru": "Участок ДО", "sr": "Deonica DO", "en": "Segment to"},
        "label.insulation": {"ru": "Тип изоляции", "sr": "Tip izolacije", "en": "Insulation type"},
        "label.conductor": {"ru": "Тип проводника", "sr": "Tip provodnika", "en": "Conductor type"},
        "label.cable": {"ru": "Марка/тип кабеля", "sr": "Oznaka/tip kabla", "en": "Cable designation"},
        "label.pi": {"ru": "Pi, Вт", "sr": "Pi, W", "en": "Pi, W"},
        "label.kj": {"ru": "Kj", "sr": "Kj", "en": "Kj"},
        "label.eta": {"ru": "η", "sr": "η", "en": "η"},
        "label.pj": {"ru": "Pj", "sr": "Pj", "en": "Pj"},
        "label.voltage": {"ru": "U, В", "sr": "U", "en": "U"},
        "label.cos_phi": {"ru": "cos φ", "sr": "cos φ", "en": "cos φ"},
        "label.length": {"ru": "Длина L, м", "sr": "Dužina L, m", "en": "Length L, m"},
        "label.area": {"ru": "Сечение, мм²", "sr": "Presek, mm²", "en": "Cross-section, mm²"},
        "label.installation": {"ru": "Способ прокладки", "sr": "Način polaganja", "en": "Installation method"},
        "label.loaded_cores": {
            "ru": "Нагруженные жилы (nž)",
            "sr": "Broj opterećenih žila (nž)",
            "en": "Loaded cores (nž)",
        },
        "label.circuits": {
            "ru": "Кабелей в группе (для S)",
            "sr": "Broj kablova u grupi (za S)",
            "en": "Cables in group (for S)",
        },
        "label.parallel": {
            "ru": "Параллельные кабели (n∥)",
            "sr": "Paralelni kablovi (n∥)",
            "en": "Parallel cables (n∥)",
        },
        "checkbox.consider_parallel": {
            "ru": "Учитывать n∥ в S",
            "sr": "Uvažavaj n∥ u S",
            "en": "Include n∥ in S",
        },
        "label.medium": {"ru": "Среда для T", "sr": "Okruženje za T", "en": "Medium for T"},
        "label.temperature": {"ru": "Температура, °C", "sr": "Temperatura, °C", "en": "Temperature, °C"},
        "label.s": {"ru": "S", "sr": "S", "en": "S"},
        "label.t": {"ru": "T", "sr": "T", "en": "T"},
        "label.in": {"ru": "In, A", "sr": "In, A", "en": "In, A"},
        "label.k": {"ru": "k", "sr": "k", "en": "k"},
        "label.drop_key": {"ru": "Ключ ΔU", "sr": "Ključ ΔU", "en": "ΔU key"},
        "label.result.pj": {"ru": "Pj, Вт", "sr": "Pj, W", "en": "Pj, W"},
        "label.result.icalc": {"ru": "Icalc [A]", "sr": "Icalc [A]", "en": "Icalc [A]"},
        "label.result.rbase": {"ru": "R_base [Ω/км]", "sr": "R_base [Ω/km]", "en": "R_base [Ω/km]"},
        "label.result.iz": {"ru": "Iz [A]", "sr": "Iz [A]", "en": "Iz [A]"},
        "label.result.s": {"ru": "S", "sr": "S", "en": "S"},
        "label.result.t": {"ru": "T", "sr": "T", "en": "T"},
        "label.result.delta": {"ru": "ΔU %", "sr": "ΔU %", "en": "ΔU %"},
        "label.result.total_delta": {"ru": "Суммарный ΔU %", "sr": "Ukupni ΔU %", "en": "Total ΔU %"},
        "label.result.limit_delta": {"ru": "Предел ΔU %", "sr": "Limit ΔU %", "en": "Limit ΔU %"},
        "label.result.ampacity": {"ru": "По току", "sr": "Po struji", "en": "By current"},
        "label.result.drop": {"ru": "По ΔU", "sr": "Po ΔU", "en": "By ΔU"},
        "label.result.in_range": {
            "ru": "Диапазон In [A]",
            "sr": "Opseg In [A]",
            "en": "In range [A]",
        },
        "label.result.i2": {"ru": "I2 [A]", "sr": "I2 [A]", "en": "I2 [A]"},
        "label.result.protection": {"ru": "Защита", "sr": "Zaštita", "en": "Protection"},
        "label.result.compatibility": {
            "ru": "Соответствие IEC",
            "sr": "IEC kompatibilnost",
            "en": "IEC compatibility",
        },
        "label.result.recommendations": {
            "ru": "Рекомендации",
            "sr": "Preporuke",
            "en": "Recommendations",
        },
        "dialog.recommendations.title": {
            "ru": "Рекомендации по подбору",
            "sr": "Preporuke za izbor",
            "en": "Selection recommendations",
        },
        "column.circuit": {"ru": "Цепь", "sr": "Strujni krug", "en": "Circuit"},
        "column.from": {"ru": "От", "sr": "OD", "en": "From"},
        "column.to": {"ru": "До", "sr": "DO", "en": "To"},
        "column.insulation": {"ru": "Изоляция", "sr": "E", "en": "E"},
        "column.conductor": {"ru": "Проводник", "sr": "F", "en": "F"},
        "column.cable": {"ru": "Кабель", "sr": "G", "en": "G"},
        "column.cores": {"ru": "nž", "sr": "nž", "en": "nž"},
        "column.n_parallel": {"ru": "n∥", "sr": "n∥", "en": "n∥"},
        "column.group_for_s": {"ru": "Кабелей в группе (S)", "sr": "Kablovi u grupi (S)", "en": "Group cables (S)"},
        "column.pi": {"ru": "Pi", "sr": "Pi", "en": "Pi"},
        "column.kj": {"ru": "Kj", "sr": "Kj", "en": "Kj"},
        "column.eta": {"ru": "η", "sr": "η", "en": "η"},
        "column.pj": {"ru": "Pj", "sr": "Pj", "en": "Pj"},
        "column.voltage": {"ru": "U", "sr": "U", "en": "U"},
        "column.cos": {"ru": "cosφ", "sr": "cosφ", "en": "cosφ"},
        "column.length": {"ru": "L", "sr": "L", "en": "L"},
        "column.area": {"ru": "Сечение", "sr": "Presek", "en": "Area"},
        "column.installation": {"ru": "Способ прокладки", "sr": "Način polaganja", "en": "Installation"},
        "column.s": {"ru": "S", "sr": "S", "en": "S"},
        "column.t": {"ru": "T", "sr": "T", "en": "T"},
        "column.in": {"ru": "In [A]", "sr": "In [A]", "en": "In [A]"},
        "column.k": {"ru": "k", "sr": "k", "en": "k"},
        "column.i2": {"ru": "I2 [A]", "sr": "I2 [A]", "en": "I2 [A]"},
        "column.icalc": {"ru": "Icalc [A]", "sr": "Icalc [A]", "en": "Icalc [A]"},
        "column.rbase": {"ru": "R_base [Ω/км]", "sr": "R_base [Ω/km]", "en": "R_base [Ω/km]"},
        "column.iz": {"ru": "Iz [A]", "sr": "Iz [A]", "en": "Iz [A]"},
        "column.drop": {"ru": "ΔU %", "sr": "ΔU %", "en": "ΔU %"},
        "column.total_drop": {"ru": "Суммарный ΔU %", "sr": "Ukupni ΔU %", "en": "Total ΔU %"},
        "column.limit_drop": {"ru": "Предел ΔU %", "sr": "Limit ΔU %", "en": "Limit ΔU %"},
        "column.ampacity": {"ru": "По току", "sr": "Po struji", "en": "By current"},
        "column.drop_status": {"ru": "По ΔU", "sr": "Po ΔU", "en": "By ΔU"},
        "column.protection": {"ru": "Защита", "sr": "Zaštita", "en": "Protection"},
        "column.key": {"ru": "Ключ", "sr": "Ključ", "en": "Key"},
        "column.compatibility": {"ru": "Соответствие IEC", "sr": "IEC kompatibilnost", "en": "IEC compatibility"},
        "column.medium": {"ru": "Среда", "sr": "Okruženje", "en": "Medium"},
        "column.limit": {"ru": "Предел ΔU %", "sr": "Limit ΔU %", "en": "Limit ΔU %"},
        "column.sigma": {"ru": "ϭ", "sr": "ϭ", "en": "ϭ"},
        "status.ok": {"ru": "OK", "sr": "OK", "en": "OK"},
        "status.fail": {"ru": "НЕТ", "sr": "NE", "en": "NO"},
        "status.na": {"ru": "Н/Д", "sr": "N/A", "en": "N/A"},
        "status.no_data": {"ru": "Нет данных", "sr": "Nema podataka", "en": "No data"},
        "warning.voltage_phase": {
            "ru": "При U=230 В трёхфазное подключение (nž=3) недопустимо. Выберите U=400 В или измените число жил.",
            "sr": "Za U=230 V trofazna veza (nž=3) nije dozvoljena. Odaberite U=400 V ili promenite broj žila.",
            "en": "At 230 V a three-phase setup (nž=3) is invalid. Choose 400 V or change the loaded cores.",
        },
        "message.select_fail": {
            "ru": "Не удалось подобрать параметры. Возможные варианты:\n",
            "sr": "Nije moguće pronaći parametre. Moguće opcije:\n",
            "en": "Unable to find suitable parameters. Possible options:\n",
        },
    }

    TOOLTIPS = {
        "label.circuit": {
            "ru": "Имя или номер цепи, используется для суммирования падений напряжения.",
            "sr": "Naziv ili broj kruga koji se koristi za sumiranje pada napona.",
            "en": "Circuit name or number used when summing voltage drop.",
        },
        "label.segment_from": {
            "ru": "Начальная точка рассматриваемой кабельной линии.",
            "sr": "Početna tačka posmatrane deonice kabla.",
            "en": "Start point of the cable segment.",
        },
        "label.segment_to": {
            "ru": "Конечная точка рассматриваемой кабельной линии.",
            "sr": "Krajnja tačka posmatrane deonice kabla.",
            "en": "End point of the cable segment.",
        },
        "label.insulation": {
            "ru": "Выберите тип изоляции кабеля согласно IEC 60364.",
            "sr": "Odaberite tip izolacije kabla prema IEC 60364.",
            "en": "Select the cable insulation type according to IEC 60364.",
        },
        "label.conductor": {
            "ru": "Материал токопроводящей жилы (медь или алюминий).",
            "sr": "Materijal provodnika (bakar ili aluminijum).",
            "en": "Conductor material (copper or aluminium).",
        },
        "label.cable": {
            "ru": "Заводская маркировка или описание кабеля.",
            "sr": "Fabricka oznaka ili opis kabla.",
            "en": "Factory designation or description of the cable.",
        },
        "label.pi": {
            "ru": "Номинальная мощность нагрузки в ваттах.",
            "sr": "Nazivna snaga opterećenja u vatima.",
            "en": "Rated load power in watts.",
        },
        "label.kj": {
            "ru": "Коэффициент спроса (одновременности) для группы потребителей.",
            "sr": "Koeficijent istovremenosti za grupu potrošača.",
            "en": "Demand (diversity) factor for the load group.",
        },
        "label.eta": {
            "ru": "КПД установки. Должен быть в диапазоне (0;1].",
            "sr": "Efikasnost sistema. Mora biti u opsegu (0;1].",
            "en": "System efficiency. Must be within (0, 1].",
        },
        "label.pj": {
            "ru": "Рассчитанная активная мощность Pi × Kj.",
            "sr": "Izračunata aktivna snaga Pi × Kj.",
            "en": "Calculated active power Pi × Kj.",
        },
        "label.voltage": {
            "ru": "Номинальное напряжение питающей сети.",
            "sr": "Nazivni napon mreže.",
            "en": "Nominal system voltage.",
        },
        "label.cos_phi": {
            "ru": "Коэффициент мощности нагрузки.",
            "sr": "Faktor snage opterećenja.",
            "en": "Load power factor.",
        },
        "label.length": {
            "ru": "Длина рассматриваемого участка кабеля в метрах.",
            "sr": "Dužina posmatrane deonice kabla u metrima.",
            "en": "Length of the analysed cable section in metres.",
        },
        "label.area": {
            "ru": "Выбранное сечение жилы кабеля.",
            "sr": "Odabrani presek provodnika.",
            "en": "Selected conductor cross-section.",
        },
        "label.installation": {
            "ru": "Метод прокладки кабеля по IEC 60364.",
            "sr": "Metod polaganja kabla prema IEC 60364.",
            "en": "Cable installation method per IEC 60364.",
        },
        "label.loaded_cores": {
            "ru": "Число нагруженных жил (2 для 1ф, 3 для 3ф систем).",
            "sr": "Broj opterećenih žila (2 za jednofazne, 3 za trofazne sisteme).",
            "en": "Number of loaded cores (2 for single-phase, 3 for three-phase).",
        },
        "label.circuits": {
            "ru": "Количество соседних кабелей в одной трассе. Влияет на коэффициент S. Ток не делится.",
            "sr": "Broj susednih kablova u istoj trasi. Utice na koeficijent S. Struja se ne deli.",
            "en": "Number of adjacent cables in one route. Affects grouping factor S. Current does not split.",
        },
        "label.parallel": {
            "ru": "Число одинаковых кабелей, подключённых параллельно к одной нагрузке. Делят ток и уменьшают падение напряжения ~ пропорционально 1/n∥.",
            "sr": "Broj identičnih kablova spojenih paralelno na jedno opterećenje. Dele struju i smanjuju pad napona približno ~1/n∥.",
            "en": "Number of identical cables connected in parallel to one load. Split the current and reduce voltage drop roughly ~1/n∥.",
        },
        "checkbox.consider_parallel": {
            "ru": "Включает параллельные кабели в расчёт коэффициента группировки S (Kn).",
            "sr": "Uključuje paralelne kablove u proračun faktora grupisanja S (Kn).",
            "en": "Includes parallel cables when evaluating grouping factor S (Kn).",
        },
        "label.medium": {
            "ru": "Среда для температурного коэффициента (воздух или грунт).",
            "sr": "Okruženje za temperaturni koeficijent (vazduh ili tlo).",
            "en": "Environment for the temperature factor (air or soil).",
        },
        "label.temperature": {
            "ru": "Фактическая температура окружающей среды.",
            "sr": "Stvarna temperatura okruženja.",
            "en": "Actual ambient temperature.",
        },
        "label.s": {
            "ru": "Коэффициент группировки Kn. Рассчитывается автоматически.",
            "sr": "Koeficijent grupisanja Kn. Računa se automatski.",
            "en": "Grouping factor Kn. Calculated automatically.",
        },
        "label.t": {
            "ru": "Температурный коэффициент Kt. Рассчитывается автоматически.",
            "sr": "Temperaturni koeficijent Kt. Računa se automatski.",
            "en": "Temperature factor Kt. Calculated automatically.",
        },
        "label.in": {
            "ru": "Номинальный ток защитного устройства.",
            "sr": "Nazivna struja zaštitnog uređaja.",
            "en": "Rated current of the protective device.",
        },
        "label.k": {
            "ru": "Коэффициент надежного срабатывания (I2/In).",
            "sr": "Koeficijent pouzdanog delovanja (I2/In).",
            "en": "Tripping reliability factor (I2/In).",
        },
        "label.drop_key": {
            "ru": "Допустимое падение напряжения по IEC 60364.",
            "sr": "Dozvoljeni pad napona prema IEC 60364.",
            "en": "Allowed voltage drop per IEC 60364.",
        },
    }

    HELP_TEXTS = {
    "ru": """Руководство по программе расчета кабелей (IEC 60364):\n\n
**Общее описание**:\n
Программа предназначена для расчета параметров кабельных линий согласно стандарту IEC 60364-5-52. Она позволяет определить допустимый ток (Iz), падение напряжения (ΔU), подобрать сечение кабеля и параметры защитного устройства, а также проверить соответствие требованиям IEC 60364-4-43. Поддерживаются медные (Cu) и алюминиевые (Al) проводники, изоляция PVC и XLPE/EPR, а также различные методы прокладки (A1, A2, B1, B2, C, D, E, F, G).\n\n
**Основные функции**:\n
- **Расчет параметров**: Вводите данные (мощность, напряжение, длина, сечение, и т.д.) для расчета тока нагрузки (Icalc), допустимого тока (Iz), падения напряжения (ΔU) и кумулятивного падения напряжения (Ukupni ΔU %).\n
- **Проверка защиты**: Проверяет соответствие номинального тока защиты (In) и тока срабатывания (I2) требованиям: Icalc ≤ In ≤ Iz и I2 ≤ 1.45 × Iz.\n
- **Автоматический подбор параметров**: Кнопка "Подобрать параметры" выбирает минимальное сечение и номинальный ток защиты (In), соответствующие требованиям по току и ΔU.\n
- **Редактирование и удаление**: Используйте кнопки "Загрузить строку в форму" и "Удалить выбранную строку" для редактирования или удаления записей в таблице результатов.\n
- **Экспорт и сохранение**: Сохраняйте проект в JSON ("Сохранить проект") и экспортируйте результаты в Excel ("Экспорт в Excel") с автофильтром и форматированием.\n
- **Локализация**: Поддержка русского, сербского и английского языков через меню "Язык".\n\n
**Коэффициенты и параметры**:\n
- **S (Kn)**: Коэффициент группировки, учитывающий взаимное нагревание кабелей. Рассчитывается автоматически на основе числа кабелей в группе. Если включен флажок "Учитывать n∥ в S", параллельные кабели (n∥) добавляются к числу кабелей в группе.\n
- **T (Kt)**: Температурный коэффициент для воздуха или грунта, выбирается автоматически из таблиц IEC 60364-5-52.\n
- **η**: КПД установки, должен быть в диапазоне (0;1]. Определяется по паспорту оборудования.\n
- **Kj**: Коэффициент спроса (одновременности) для группы потребителей.\n
- **ΔU**: Допустимое падение напряжения, определяется ключом (UIDM: 5%, SVDM: 3%, SVTS: 5%, UITS: 8%).\n
- **cos φ**: Коэффициент мощности нагрузки, должен быть в диапазоне (0;1].\n
- **In, k**: Параметры защитного устройства: номинальный ток (In) и коэффициент срабатывания (I2/In).\n
- **ϭ**: Специфическая проводимость проводника (м/Ω·мм²), рассчитывается как 1/ρ, где ρ — удельное сопротивление при 20°C.\n
- **n∥**: Число параллельных кабелей, делящих ток нагрузки (Icalc/n∥) и уменьшающих сопротивление (R/n∥, X/n∥).\n\n
**Методы прокладки (IEC 60364-5-52)**:\n
- **A1**: Многожильные кабели в трубе, проложенной в теплоизолированной стене (низкая теплопроводность).\n
- **A2**: Многожильные кабели в трубе, проложенной в обычной стене или кладке (высокая теплопроводность).\n
- **B1**: Одножильные кабели в трубе, проложенной в теплоизолированной стене.\n
- **B2**: Одножильные кабели в трубе, проложенной в обычной стене или кладке.\n
- **C**: Многожильные или одножильные кабели на поверхности (например, на стене или потолке) без труб.\n
- **D**: Кабели, проложенные в грунте (в траншее или трубе под землёй).\n
- **E**: Многожильные кабели в свободном воздухе (на открытой трассе, без касания поверхностей).\n
- **F**: Одножильные кабели в свободном воздухе, с расстоянием между кабелями не менее диаметра.\n
- **G**: Одножильные кабели в свободном воздухе, с минимальным расстоянием или касанием друг друга.\n\n
**Инструкции по использованию**:\n
1. Заполните поля ввода: укажите цепь, мощность (Pi), Kj, η, напряжение (U), cos φ, длину (L), сечение, метод прокладки, число жил (nž), и т.д.\n
2. Убедитесь, что U=230 В используется только с nж=2 (однофазная система), а U=400 В — с nж=3 (трёхфазная).\n
3. Выберите метод прокладки в соответствии с условиями установки (например, D для грунта, C для поверхностной прокладки).\n
4. Используйте "Подобрать параметры" для автоматического выбора сечения и In.\n
5. Нажмите "Рассчитать и добавить строку", чтобы сохранить результат в таблице.\n
6. Проверьте "Совместимость IEC": должно быть "OK" для соответствия стандарту.\n
7. При необходимости отредактируйте строку через "Загрузить строку в форму" или удалите через "Удалить выбранную строку".\n
8. Сохраните проект или экспортируйте результаты в Excel.\n\n
**Замечания**:\n
- Если температура выходит за пределы таблиц IEC, программа выдаст предупреждение.\n
- При несоответствии по току (Icalc > Iz) или ΔU программа предложит рекомендации.\n
- Проверяйте корректность ввода, чтобы избежать ошибок в расчетах.\n
- Выбор метода прокладки существенно влияет на допустимый ток (Iz). Используйте справку для выбора подходящего метода.\n""",
    "sr": """Uputstvo za program za proračun kablova (IEC 60364):\n\n
**Opšti opis**:\n
Program je namenjen za proračun parametara kablovskih linija prema standardu IEC 60364-5-52. Omogućava određivanje dopustive struje (Iz), pada napona (ΔU), odabir preseka kabla i parametara zaštitnog uređaja, kao i proveru usaglašenosti sa zahtevima IEC 60364-4-43. Podržava bakarne (Cu) i aluminijumske (Al) provodnike, izolaciju PVC i XLPE/EPR, i različite metode polaganja (A1, A2, B1, B2, C, D, E, F, G).\n\n
**Glavne funkcije**:\n
- **Proračun parametara**: Unesite podatke (snaga, napon, dužina, presek, itd.) za izračunavanje struje opterećenja (Icalc), dopustive struje (Iz), pada napona (ΔU) i ukupnog pada napona (Ukupni ΔU %).\n
- **Provera zaštite**: Proverava usaglašenost nazivne struje zaštite (In) i struje okidanja (I2) sa uslovima: Icalc ≤ In ≤ Iz i I2 ≤ 1.45 × Iz.\n
- **Automatski odabir parametara**: Dugme "Odaberi parametre" bira minimalni presek i nazivnu struju zaštite (In) koji zadovoljavaju zahteve za struju i ΔU.\n
- **Uređivanje i brisanje**: Koristite dugmad "Učitaj red u formu" i "Obriši izabrani red" za uređivanje ili uklanjanje unosa u tabeli rezultata.\n
- **Izvoz i čuvanje**: Sačuvajte projekat u JSON formatu ("Sačuvaj projekat") i izvezite rezultate u Excel ("Izvoz u Excel") sa automatskim filterima i formatiranjem.\n
- **Lokalizacija**: Podrška za srpski, ruski i engleski jezik putem menija "Jezik".\n\n
**Koeficijenti i parametri**:\n
- **S (Kn)**: Koeficijent grupisanja koji uzima u obzir međusobno zagrevanje kablova. Automatski se računa na osnovu broja kablova u grupi. Ako je označeno polje "Uvažavaj n∥ u S", paralelni kablovi (n∥) se dodaju broju kablova u grupi.\n
- **T (Kt)**: Temperaturni koeficijent za vazduh ili tlo, automatski se bira iz tabela IEC 60364-5-52.\n
- **η**: Efikasnost postrojenja, mora biti u opsegu (0;1]. Određuje se prema podacima proizvođača.\n
- **Kj**: Koeficijent istovremenosti za grupu potrošača.\n
- **ΔU**: Dozvoljeni pad napona, određen ključem (UIDM: 5%, SVDM: 3%, SVTS: 5%, UITS: 8%).\n
- **cos φ**: Faktor snage opterećenja, mora biti u opsegu (0;1].\n
- **In, k**: Parametri zaštitnog uređaja: nazivna struja (In) i koeficijent okidanja (I2/In).\n
- **ϭ**: Specifična provodnost provodnika (m/Ω·mm²), izračunava se kao 1/ρ, gde je ρ specifično otpornost na 20°C.\n
- **n∥**: Broj paralelnih kablova koji dele struju opterećenja (Icalc/n∥) i smanjuju otpor (R/n∥, X/n∥).\n\n
**Metode polaganja (IEC 60364-5-52)**:\n
- **A1**: Višežilni kablovi u cevi, položeni u toplotno izolovani zid (niska toplotna provodljivost).\n
- **A2**: Višežilni kablovi u cevi, položeni u običan zid ili zidanu konstrukciju (visoka toplotna provodljivost).\n
- **B1**: Jednožilni kablovi u cevi, položeni u toplotno izolovani zid.\n
- **B2**: Jednožilni kablovi u cevi, položeni u običan zid ili zidanu konstrukciju.\n
- **C**: Višežilni ili jednožilni kablovi na površini (npr. na zidu ili plafonu) bez cevi.\n
- **D**: Kablovi položeni u tlo (u rovu ili cevi ispod zemlje).\n
- **E**: Višežilni kablovi na slobodnom vazduhu (na otvorenoj trasi, bez dodira sa površinama).\n
- **F**: Jednožilni kablovi na slobodnom vazduhu, sa razmakom između kablova najmanje jednog prečnika.\n
- **G**: Jednožilni kablovi na slobodnom vazduhu, sa minimalnim razmakom ili u dodiru jedan s drugim.\n\n
**Uputstvo za korišćenje**:\n
1. Popunite polja za unos: unesite krug, snagu (Pi), Kj, η, napon (U), cos φ, dužinu (L), presek, metod polaganja, broj žila (nž), itd.\n
2. Uverite se da je U=230 V korišćeno samo sa nž=2 (jednofazni sistem), a U=400 V sa nž=3 (trofazni sistem).\n
3. Odaberite metod polaganja u skladu sa uslovima instalacije (npr. D za tlo, C za površinsko polaganje).\n
4. Koristite "Odaberi parametre" za automatski izbor preseka i In.\n
5. Kliknite na "Izračunaj i dodaj red" da biste sačuvali rezultat u tabelu.\n
6. Proverite "IEC kompatibilnost": treba da bude "OK" za usaglašenost sa standardom.\n
7. Po potrebi uredite red pomoću "Učitaj red u formu" ili obrišite pomoću "Obriši izabrani red".\n
8. Sačuvajte projekat ili izvezite rezultate u Excel.\n\n
**Napomene**:\n
- Ako temperatura izlazi izvan opsega tabela IEC, program će izdati upozorenje.\n
- Ako postoji neusaglašenost po struji (Icalc > Iz) ili ΔU, program će predložiti preporuke.\n
- Proverite ispravnost unosa kako biste izbegli greške u proračunima.\n
- Izbor metode polaganja značajno utiče na dopustivu struju (Iz). Koristite pomoć za odabir odgovarajuće metode.\n""",
    "en": """Guide to the Cable Calculation Program (IEC 60364):\n\n
**Overview**:\n
This program is designed to calculate cable parameters according to IEC 60364-5-52. It determines the permissible current (Iz), voltage drop (ΔU), selects cable cross-sections and protective device parameters, and verifies compliance with IEC 60364-4-43. It supports copper (Cu) and aluminum (Al) conductors, PVC and XLPE/EPR insulation, and various installation methods (A1, A2, B1, B2, C, D, E, F, G).\n\n
**Main Features**:\n
- **Parameter Calculation**: Enter data (power, voltage, length, cross-section, etc.) to calculate load current (Icalc), permissible current (Iz), voltage drop (ΔU), and cumulative voltage drop (Ukupni ΔU %).\n
- **Protection Check**: Verifies compliance of the protective devices rated current (In) and tripping current (I2) with conditions: Icalc ≤ In ≤ Iz and I2 ≤ 1.45 × Iz.\n
- **Automatic Parameter Selection**: The "Select parameters" button chooses the minimum cross-section and rated current (In) that meet current and ΔU requirements.\n
- **Editing and Deletion**: Use the "Load row into form" and "Remove selected row" buttons to edit or delete entries in the results table.\n
- **Export and Save**: Save the project in JSON ("Save project") and export results to Excel ("Export to Excel") with autofilter and formatting.\n
- **Localization**: Supports Russian, Serbian, and English languages via the "Language" menu.\n\n
**Coefficients and Parameters**:\n
- **S (Kn)**: Grouping factor accounting for mutual heating of cables. Calculated automatically based on the number of cables in a group. If the "Include n∥ in S" checkbox is enabled, parallel cables (n∥) are added to the group count.\n
- **T (Kt)**: Temperature factor for air or soil, automatically selected from IEC 60364-5-52 tables.\n
- **η**: Installation efficiency, must be in the range (0;1]. Determined from equipment specifications.\n
- **Kj**: Demand (diversity) factor for a group of loads.\n
- **ΔU**: Permissible voltage drop, defined by the key (UIDM: 5%, SVDM: 3%, SVTS: 5%, UITS: 8%).\n
- **cos φ**: Load power factor, must be in the range (0;1].\n
- **In, k**: Protective device parameters: rated current (In) and tripping ratio (I2/In).\n
- **ϭ**: Specific conductivity of the conductor (m/Ω·mm²), calculated as 1/ρ, where ρ is the resistivity at 20°C.\n
- **n∥**: Number of parallel cables sharing the load current (Icalc/n∥) and reducing resistance (R/n∥, X/n∥).\n\n
**Installation Methods (IEC 60364-5-52)**:\n
- **A1**: Multicore cables in conduit within a thermally insulated wall (low thermal conductivity).\n
- **A2**: Multicore cables in conduit within a normal wall or masonry (high thermal conductivity).\n
- **B1**: Single-core cables in conduit within a thermally insulated wall.\n
- **B2**: Single-core cables in conduit within a normal wall or masonry.\n
- **C**: Multicore or single-core cables on a surface (e.g., on a wall or ceiling) without conduit.\n
- **D**: Cables buried in the ground (in a trench or conduit underground).\n
- **E**: Multicore cables in free air (on an open tray, not touching surfaces).\n
- **F**: Single-core cables in free air, spaced at least one cable diameter apart.\n
- **G**: Single-core cables in free air, with minimal spacing or touching each other.\n\n
**Usage Instructions**:\n
1. Fill in the input fields: specify the circuit, power (Pi), Kj, η, voltage (U), cos φ, length (L), cross-section, installation method, number of cores (nž), etc.\n
2. Ensure U=230 V is used only with nž=2 (single-phase system), and U=400 V with nž=3 (three-phase system).\n
3. Select the installation method based on the installation conditions (e.g., D for ground, C for surface mounting).\n
4. Use "Select parameters" to automatically choose cross-section and In.\n
5. Click "Calculate and add row" to save the result to the table.\n
6. Check "IEC compatibility": it should be "OK" for standard compliance.\n
7. Edit a row using "Load row into form" or delete it with "Remove selected row".\n
8. Save the project or export results to Excel.\n\n
**Notes**:\n
- If the temperature is outside IEC table ranges, a warning will be displayed.\n
- If there is non-compliance in current (Icalc > Iz) or ΔU, recommendations will be provided.\n
- Verify input correctness to avoid calculation errors.\n
- The choice of installation method significantly affects the permissible current (Iz). Use the help to select the appropriate method.\n""",
}

    LABEL_KEY_MAP = {
        "Strujni krug": "label.circuit",
        "Deonica OD": "label.segment_from",
        "Deonica DO": "label.segment_to",
        "Tip-IZOLACIJE": "label.insulation",
        "Tip-PROVODNIKA": "label.conductor",
        "Oznaka-tip-KABLA": "label.cable",
        "Pi, W": "label.pi",
        "Kj": "label.kj",
        "η": "label.eta",
        "Pj": "label.pj",
        "U": "label.voltage",
        "cos φ": "label.cos_phi",
        "Dužina L, m": "label.length",
        "Presek, mm²": "label.area",
        "Način polaganja": "label.installation",
        "Нагруженные жилы (nž)": "label.loaded_cores",
        "Кабелей в группе (для S)": "label.circuits",
        "Параллельные кабели (n∥)": "label.parallel",
        "Учитывать n∥ в S": "checkbox.consider_parallel",
        "Среда для Т": "label.medium",
        "Температура, °C": "label.temperature",
        "S": "label.s",
        "T": "label.t",
        "In, A": "label.in",
        "k": "label.k",
        "Ключ ΔU": "label.drop_key",
    }

    RESULT_LABEL_KEY_MAP = {
        "Pj, W": "label.result.pj",
        "Icalc [A]": "label.result.icalc",
        "R_base [Ω/km]": "label.result.rbase",
        "Iz [A]": "label.result.iz",
        "S": "label.result.s",
        "T": "label.result.t",
        "ΔU %": "label.result.delta",
        "Ukupni ΔU %": "label.result.total_delta",
        "Limit ΔU %": "label.result.limit_delta",
        "По току": "label.result.ampacity",
        "По ΔU": "label.result.drop",
        "Диапазон In [A]": "label.result.in_range",
        "I2 [A]": "label.result.i2",
        "Защита": "label.result.protection",
        "Совместимость IEC": "label.result.compatibility",
        "Рекомендации": "label.result.recommendations",
    }

    TREE_COLUMN_KEYS = {
        "Strujni krug": "column.circuit",
        "OD": "column.from",
        "DO": "column.to",
        "E": "column.insulation",
        "F": "column.conductor",
        "G": "column.cable",
        "nž": "column.cores",
        "n∥": "column.n_parallel",
        "Кабелей в группе (S)": "column.group_for_s",
        "Pi": "column.pi",
        "Kj": "column.kj",
        "η": "column.eta",
        "Pj": "column.pj",
        "U": "column.voltage",
        "cosφ": "column.cos",
        "L": "column.length",
        "Presek": "column.area",
        "Način polaganja": "column.installation",
        "S": "column.s",
        "T": "column.t",
        "In [A]": "column.in",
        "k": "column.k",
        "I2 [A]": "column.i2",
        "Icalc [A]": "column.icalc",
        "R_base [Ω/km]": "column.rbase",
        "ϭ": "column.sigma",
        "Iz [A]": "column.iz",
        "ΔU %": "column.drop",
        "Ukupni ΔU %": "column.total_drop",
        "Limit ΔU %": "column.limit_drop",
        "По току": "column.ampacity",
        "По ΔU": "column.drop_status",
        "Защита": "column.protection",
        "Ключ": "column.key",
        "Совместимость IEC": "column.compatibility",
    }

    INSULATION_OPTIONS = [
        "PVC (70°C)",
        "XLPE/EPR (90°C)",
    ]
    INSULATION_META = {
        "PVC (70°C)": {"key": "PVC", "theta": 70},
        "XLPE/EPR (90°C)": {"key": "XLPE", "theta": 90},
    }

    CONDUCTOR_TYPES = ["Cu", "Al"]
    VOLTAGE_LEVELS = ["400", "230"]
    TEMPERATURE_MEDIA = {
        "air": {"ru": "Воздух", "sr": "Vazduh", "en": "Air"},
        "soil": {"ru": "Грунт", "sr": "Tlo", "en": "Soil"},
    }
    INSTALLATION_METHODS = ["A1", "A2", "B1", "B2", "C", "D", "E", "F", "G"]
    STANDARD_SECTIONS = [
        1.5,
        2.5,
        4,
        6,
        10,
        16,
        25,
        35,
        50,
        70,
        95,
        120,
        150,
        185,
        240,
        300,
        400,
        500,
    ]
    METHOD_PREFERENCE = ["D", "E", "F", "C", "B2", "B1", "A2", "A1", "G"]
    STANDARD_CROSS_SECTIONS = [
        "",
        "0.5",
        "0.75",
        "1",
        "1.5",
        "2.5",
        "4",
        "6",
        "10",
        "16",
        "25",
        "35",
        "50",
        "70",
        "95",
        "120",
        "150",
        "185",
        "240",
        "300",
        "400",
        "500",
        "630",
    ]
    STANDARD_BREAKER_RATINGS = [
        "",
        "6",
        "10",
        "13",
        "16",
        "20",
        "25",
        "32",
        "40",
        "50",
        "63",
        "80",
        "100",
        "125",
        "160",
        "200",
        "250",
        "315",
        "400",
        "500",
        "630",
    ]
    DROP_LIMIT_KEYS = {
        "UIDM": 5.0,
        "SVDM": 3.0,
        "SVTS": 5.0,
        "UITS": 8.0,
    }

    RESISTIVITY_20 = {"Cu": 0.017241, "Al": 0.028264}
    TEMP_COEFF = {"Cu": 0.00393, "Al": 0.00403}
    REACTANCE_PER_KM = {
        "default": 0.08,
        "D": 0.09,
        ("D", "≤95"): 0.09,
        ("D", "≤240"): 0.085,
        ("D", ">240"): 0.08,
        ("C", "≤95"): 0.08,
        ("C", "≤240"): 0.077,
        ("C", ">240"): 0.074,
        ("F", "≤95"): 0.08,
        ("F", "≤240"): 0.078,
        ("F", ">240"): 0.075,
    }

    AMPACITY_BASE = {
        "A1": {
            1.5: 17.0,
            2.5: 23.0,
            4.0: 30.0,
            6.0: 38.0,
            10.0: 52.0,
            16.0: 69.0,
            25.0: 89.0,
            35.0: 110.0,
            50.0: 132.0,
            70.0: 168.0,
            95.0: 202.0,
            120.0: 231.0,
            150.0: 265.0,
            185.0: 301.0,
            240.0: 352.0,
            300.0: 395.0,
            400.0: 450.0,
            500.0: 500.0,
        },
        "A2": {
            1.5: 18.0,
            2.5: 24.0,
            4.0: 32.0,
            6.0: 40.0,
            10.0: 54.0,
            16.0: 72.0,
            25.0: 94.0,
            35.0: 117.0,
            50.0: 141.0,
            70.0: 179.0,
            95.0: 216.0,
            120.0: 249.0,
            150.0: 285.0,
            185.0: 324.0,
            240.0: 380.0,
            300.0: 425.0,
            400.0: 490.0,
            500.0: 550.0,
        },
        "B1": {
            1.5: 19.0,
            2.5: 26.0,
            4.0: 34.0,
            6.0: 44.0,
            10.0: 60.0,
            16.0: 76.0,
            25.0: 101.0,
            35.0: 123.0,
            50.0: 146.0,
            70.0: 185.0,
            95.0: 225.0,
            120.0: 260.0,
            150.0: 300.0,
            185.0: 344.0,
            240.0: 404.0,
            300.0: 460.0,
            400.0: 525.0,
            500.0: 590.0,
        },
        "B2": {
            1.5: 21.0,
            2.5: 28.0,
            4.0: 37.0,
            6.0: 48.0,
            10.0: 65.0,
            16.0: 84.0,
            25.0: 110.0,
            35.0: 135.0,
            50.0: 162.0,
            70.0: 204.0,
            95.0: 244.0,
            120.0: 281.0,
            150.0: 323.0,
            185.0: 370.0,
            240.0: 435.0,
            300.0: 495.0,
            400.0: 565.0,
            500.0: 630.0,
        },
        "C": {
            1.5: 20.0,
            2.5: 27.0,
            4.0: 36.0,
            6.0: 46.0,
            10.0: 61.0,
            16.0: 80.0,
            25.0: 104.0,
            35.0: 125.0,
            50.0: 150.0,
            70.0: 192.0,
            95.0: 232.0,
            120.0: 269.0,
            150.0: 309.0,
            185.0: 355.0,
            240.0: 415.0,
            300.0: 480.0,
            400.0: 555.0,
            500.0: 625.0,
        },
        "D": {
            1.5: 25.0,
            2.5: 33.0,
            4.0: 43.0,
            6.0: 55.0,
            10.0: 75.0,
            16.0: 95.0,
            25.0: 115.0,
            35.0: 140.0,
            50.0: 170.0,
            70.0: 215.0,
            95.0: 260.0,
            120.0: 300.0,
            150.0: 340.0,
            185.0: 385.0,
            240.0: 455.0,
            300.0: 520.0,
            400.0: 600.0,
            500.0: 680.0,
        },
        "E": {
            1.5: 25.0,
            2.5: 33.0,
            4.0: 44.0,
            6.0: 56.0,
            10.0: 75.0,
            16.0: 100.0,
            25.0: 127.0,
            35.0: 154.0,
            50.0: 185.0,
            70.0: 229.0,
            95.0: 273.0,
            120.0: 312.0,
            150.0: 356.0,
            185.0: 402.0,
            240.0: 467.0,
            300.0: 535.0,
            400.0: 615.0,
            500.0: 700.0,
        },
        "F": {
            1.5: 29.0,
            2.5: 38.0,
            4.0: 51.0,
            6.0: 65.0,
            10.0: 89.0,
            16.0: 119.0,
            25.0: 156.0,
            35.0: 191.0,
            50.0: 231.0,
            70.0: 295.0,
            95.0: 357.0,
            120.0: 412.0,
            150.0: 476.0,
            185.0: 546.0,
            240.0: 640.0,
            300.0: 720.0,
            400.0: 820.0,
            500.0: 920.0,
        },
        "G": {
            1.5: 27.0,
            2.5: 36.0,
            4.0: 48.0,
            6.0: 61.0,
            10.0: 83.0,
            16.0: 111.0,
            25.0: 146.0,
            35.0: 179.0,
            50.0: 217.0,
            70.0: 276.0,
            95.0: 334.0,
            120.0: 386.0,
            150.0: 446.0,
            185.0: 511.0,
            240.0: 598.0,
            300.0: 670.0,
            400.0: 760.0,
            500.0: 850.0,
        },
    }

    AMPACITY_INSULATION_FACTORS = {
        "PVC": {
            "Cu": {
                "A1": 1.0,
                "A2": 1.0,
                "B1": 1.0,
                "B2": 1.0,
                "C": 1.0,
                "D": 1.0,
                "E": 1.0,
                "F": 1.0,
                "G": 1.0,
            },
            "Al": {
                "A1": 0.76,
                "A2": 0.76,
                "B1": 0.76,
                "B2": 0.76,
                "C": 0.76,
                "D": 0.76,
                "E": 0.76,
                "F": 0.76,
                "G": 0.76,
            },
        },
        "XLPE": {
            "Cu": {
                "A1": 1.2,
                "A2": 1.2,
                "B1": 1.21,
                "B2": 1.18,
                "C": 1.27,
                "D": 1.2,
                "E": 1.23,
                "F": 1.24,
                "G": 1.22,
            },
            "Al": {
                "A1": 0.92,
                "A2": 0.92,
                "B1": 0.92,
                "B2": 0.9,
                "C": 0.96,
                "D": 0.92,
                "E": 0.94,
                "F": 0.95,
                "G": 0.93,
            },
        },
    }

    AMPACITY_LOADED_FACTORS = {
        "A1": {2: 1.18, 3: 1.0},
        "A2": {2: 1.18, 3: 1.0},
        "B1": {2: 1.17, 3: 1.0},
        "B2": {2: 1.15, 3: 1.0},
        "C": {2: 1.17, 3: 1.0},
        "D": {2: 1.15, 3: 1.0},
        "E": {2: 1.16, 3: 1.0},
        "F": {2: 1.13, 3: 1.0},
        "G": {2: 1.13, 3: 1.0},
    }

    GROUPING_FACTORS = {
        1: 1.0,
        2: 0.8,
        3: 0.7,
        4: 0.65,
        5: 0.6,
        6: 0.57,
        7: 0.54,
        8: 0.52,
        9: 0.5,
        10: 0.48,
        11: 0.47,
        12: 0.46,
        13: 0.45,
        14: 0.44,
        15: 0.43,
        16: 0.42,
        17: 0.41,
        18: 0.4,
        19: 0.39,
        20: 0.38,
    }

    KT_V_TABLE = {
        "PVC": {
            10: 1.22,
            15: 1.17,
            20: 1.12,
            25: 1.06,
            30: 1.0,
            35: 0.94,
            40: 0.87,
            45: 0.79,
            50: 0.71,
            55: 0.61,
            60: 0.5,
            65: 0.35,
            70: 0.2,
        },
        "XLPE": {
            10: 1.15,
            15: 1.12,
            20: 1.08,
            25: 1.04,
            30: 1.0,
            35: 0.96,
            40: 0.91,
            45: 0.87,
            50: 0.82,
            55: 0.76,
            60: 0.71,
            65: 0.65,
            70: 0.58,
            75: 0.5,
            80: 0.41,
            85: 0.29,
            90: 0.2,
        },
    }

    KT_Z_TABLE = {
        "PVC": {
            10: 1.1,
            15: 1.06,
            20: 1.0,
            25: 0.95,
            30: 0.9,
            35: 0.86,
            40: 0.82,
            45: 0.78,
            50: 0.75,
            55: 0.72,
            60: 0.7,
        },
        "XLPE": {
            10: 1.06,
            15: 1.03,
            20: 1.0,
            25: 0.97,
            30: 0.94,
            35: 0.91,
            40: 0.88,
            45: 0.85,
            50: 0.82,
            55: 0.79,
            60: 0.76,
        },
    }

    # Neutralize embedded dictionaries: use external data only
    TRANSLATIONS: dict[str, dict[str, str]] = {}
    TOOLTIPS: dict[str, dict[str, str]] = {}
    HELP_TEXTS: dict[str, str] = {}
    INSULATION_OPTIONS: list[str] = []
    INSULATION_META: dict[str, dict[str, typing.Any]] = {}
    CONDUCTOR_TYPES: list[str] = []
    VOLTAGE_LEVELS: list[str] = []
    TEMPERATURE_MEDIA: dict[str, dict[str, str]] = {}
    INSTALLATION_METHODS: list[str] = []
    STANDARD_SECTIONS: list[float] = []
    METHOD_PREFERENCE: list[str] = []
    STANDARD_CROSS_SECTIONS: list[str] = []
    STANDARD_BREAKER_RATINGS: list[str] = []
    DROP_LIMIT_KEYS: dict[str, float] = {}
    RESISTIVITY_20: dict[str, float] = {}
    TEMP_COEFF: dict[str, float] = {}
    AMPACITY_BASE: dict[str, dict[float, float]] = {}
    AMPACITY_INSULATION_FACTORS: dict[str, dict[str, dict[str, float]]] = {}
    AMPACITY_LOADED_FACTORS: dict[str, dict[int, float]] = {}
    GROUPING_FACTORS: dict[int, float] = {}
    KT_V_TABLE: dict[str, dict[int, float]] = {}
    KT_Z_TABLE: dict[str, dict[int, float]] = {}
    REACTANCE_DATA: dict[str, typing.Any] = {}

    TREE_COLUMNS = (
        "Strujni krug",
        "OD",
        "DO",
        "E",
        "F",
        "G",
        "nž",
        "n∥",
        "Кабелей в группе (S)",
        "Pi",
        "Kj",
        "η",
        "Pj",
        "U",
        "cosφ",
        "L",
        "Presek",
        "Način polaganja",
        "S",
        "T",
        "In [A]",
        "k",
        "I2 [A]",
        "Icalc [A]",
        "R_base [Ω/km]",
        "ϭ",
        "Iz [A]",
        "ΔU %",
        "Ukupni ΔU %",
        "Limit ΔU %",
        "По току",
        "По ΔU",
        "Защита",
        "Ключ",
        "Совместимость IEC",
    )

    def __init__(self) -> None:
        super().__init__()
        self.geometry(self.WINDOW_GEOMETRY)

        self._language = tk.StringVar(value=self.DEFAULT_LANGUAGE)
        self._language.trace_add("write", self._on_language_change)

        self._text_bindings: list[tuple[typing.Callable[[str], None], str]] = []
        self._menu_text_bindings: list[tuple[tk.Menu, int, str]] = []
        self._tree_heading_bindings: list[tuple[str, str]] = []

        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        default_fg = self.style.lookup("TLabel", "foreground") or "#202020"
        default_bg = self.style.lookup("TLabel", "background") or self.cget("background")
        self.style.configure("ResultKey.TLabel", foreground=default_fg, background=default_bg)
        self.style.configure("ResultValue.TLabel", foreground=default_fg, background=default_bg)
        self.style.configure(
            "ResultAlert.TLabel",
            foreground="#b00020",
            background="#ffe6e6",
        )
        self.style.map("ResultAlert.TLabel", background=[("!disabled", "#ffe6e6")])
        self.style.configure("Invalid.TEntry", fieldbackground="#ffe6e6")
        self.style.map("Invalid.TEntry", fieldbackground=[("!disabled", "#ffe6e6")])
        self.style.configure("Invalid.TCombobox", fieldbackground="#ffe6e6")
        self.style.map(
            "Invalid.TCombobox",
            fieldbackground=[("readonly", "#ffe6e6"), ("!disabled", "#ffe6e6")],
        )

        self._form_values: dict[str, tk.Variable] = {}
        self._input_widgets: dict[str, ttk.Widget] = {}
        self._input_styles: dict[str, str] = {}
        self._combobox_values: dict[str, list[str]] = {}
        self._intermediate_vars: dict[str, tk.StringVar] = {}
        self._intermediate_labels: dict[str, ttk.Label] = {}
        self._table_data: list[dict[str, str]] = []
        self._last_temperature_warning: tuple[str, str, float] | None = None
        self._temperature_editing = False
        self._tooltips: list["Tooltip"] = []
        if self.TEMPERATURE_MEDIA:
            self._medium_selected_key = (
                self.DEFAULT_MEDIUM
                if self.DEFAULT_MEDIUM in self.TEMPERATURE_MEDIA
                else next(iter(self.TEMPERATURE_MEDIA))
            )
        else:
            self._medium_selected_key = self.DEFAULT_MEDIUM
        self._medium_combobox: ttk.Combobox | None = None
        self._help_text_widget: tk.Text | None = None
        self._notebook: ttk.Notebook | None = None
        self._tabs: dict[str, ttk.Frame] = {}
        self._last_icalc: float | None = None
        self._consider_parallel_in_s = tk.BooleanVar(value=False)
        self._voltage_phase_warning_shown = False

        # Load external resources (translations, tooltips, numeric tables)
        self._load_external_resources()

        self._build_menu()
        self._build_layout()

        self.title(self._(self.WINDOW_TITLE_KEY))
        self._apply_language()

    def _(self, key: str) -> str:
        translations = self.TRANSLATIONS.get(key)
        if not translations:
            return key
        language = self._language.get()
        if language in translations:
            return translations[language]
        default_value = translations.get(self.DEFAULT_LANGUAGE)
        if default_value is not None:
            return default_value
        return next(iter(translations.values()))

    def _bind_text(self, setter: typing.Callable[[str], None], key: str) -> None:
        self._text_bindings.append((setter, key))
        setter(self._(key))

    def _register_menu_text(self, menu: tk.Menu, index: int, key: str) -> None:
        self._menu_text_bindings.append((menu, index, key))
        menu.entryconfigure(index, label=self._(key))

    def _register_tree_heading(self, column_id: str, key: str) -> None:
        self._tree_heading_bindings.append((column_id, key))
        self.tree.heading(column_id, text=self._(key))

    def _register_notebook_tab(self, tab: ttk.Frame, key: str) -> None:
        if not self._notebook:
            return

        def setter(value: str, tab_ref: ttk.Frame = tab) -> None:
            if self._notebook:
                self._notebook.tab(tab_ref, text=value)

        self._bind_text(setter, key)

    def _on_language_change(self, *_: object) -> None:
        self._apply_language()

    def _apply_language(self) -> None:
        self.title(self._(self.WINDOW_TITLE_KEY))

        for setter, key in self._text_bindings:
            setter(self._(key))

        for menu, index, key in self._menu_text_bindings:
            try:
                menu.entryconfigure(index, label=self._(key))
            except tk.TclError:
                continue

        if hasattr(self, "tree"):
            for column_id, key in self._tree_heading_bindings:
                try:
                    self.tree.heading(column_id, text=self._(key))
                except tk.TclError:
                    continue

        self._update_medium_options()
        self._update_help_text()

    def _resource_path(self, *parts: str) -> str:
        base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, "data", *parts)

    def _load_json_file(self, rel_path: str) -> typing.Any:
        try:
            with open(self._resource_path(rel_path), "r", encoding="utf-8") as f:
                return json.load(f)
        except OSError:
            return None
        except json.JSONDecodeError:
            logging.error("JSON parse error for %s", rel_path)
            return None

    def _load_external_resources(self) -> None:
        # Translations
        translations = self._load_json_file("translations.json")
        if isinstance(translations, dict):
            self.TRANSLATIONS = translations
        tooltips = self._load_json_file("tooltips.json")
        if isinstance(tooltips, dict):
            self.TOOLTIPS = tooltips
        # Help texts
        helps: dict[str, str] = {}
        for code in ("ru", "sr", "en"):
            try:
                with open(self._resource_path("help", f"{code}.txt"), "r", encoding="utf-8") as f:
                    helps[code] = f.read()
            except OSError:
                continue
        if helps:
            self.HELP_TEXTS = helps

        # Tables and numeric data
        tables = self._load_json_file("tables.json")
        if not isinstance(tables, dict):
            return
        self.INSULATION_OPTIONS = tables.get("INSULATION_OPTIONS", [])
        self.INSULATION_META = tables.get("INSULATION_META", {})
        self.CONDUCTOR_TYPES = tables.get("CONDUCTOR_TYPES", [])
        self.VOLTAGE_LEVELS = tables.get("VOLTAGE_LEVELS", [])
        self.TEMPERATURE_MEDIA = tables.get("TEMPERATURE_MEDIA", {})
        self.INSTALLATION_METHODS = tables.get("INSTALLATION_METHODS", [])
        self.STANDARD_SECTIONS = tables.get("STANDARD_SECTIONS", [])
        self.METHOD_PREFERENCE = tables.get("METHOD_PREFERENCE", [])
        self.STANDARD_CROSS_SECTIONS = tables.get("STANDARD_CROSS_SECTIONS", [])
        self.STANDARD_BREAKER_RATINGS = tables.get("STANDARD_BREAKER_RATINGS", [])
        self.DROP_LIMIT_KEYS = tables.get("DROP_LIMIT_KEYS", {})
        self.RESISTIVITY_20 = tables.get("RESISTIVITY_20", {})
        self.TEMP_COEFF = tables.get("TEMP_COEFF", {})
        self.REACTANCE_DATA = tables.get("REACTANCE", {})

        # Convert dicts that require numeric keys
        amp_base_raw = tables.get("AMPACITY_BASE", {})
        amp_base: dict[str, dict[float, float]] = {}
        for m, inner in amp_base_raw.items():
            try:
                amp_base[m] = {float(k): float(v) for k, v in inner.items()}
            except Exception:
                continue
        if amp_base:
            self.AMPACITY_BASE = amp_base

        self.AMPACITY_INSULATION_FACTORS = tables.get("AMPACITY_INSULATION_FACTORS", {})

        loaded_raw = tables.get("AMPACITY_LOADED_FACTORS", {})
        loaded: dict[str, dict[int, float]] = {}
        for m, inner in loaded_raw.items():
            try:
                loaded[m] = {int(k): float(v) for k, v in inner.items()}
            except Exception:
                continue
        if loaded:
            self.AMPACITY_LOADED_FACTORS = loaded

        grouping_raw = tables.get("GROUPING_FACTORS", {})
        try:
            self.GROUPING_FACTORS = {int(k): float(v) for k, v in grouping_raw.items()}
        except Exception:
            pass

        ktv_raw = tables.get("KT_V_TABLE", {})
        ktv: dict[str, dict[int, float]] = {}
        for ins, inner in ktv_raw.items():
            try:
                ktv[ins] = {int(k): float(v) for k, v in inner.items()}
            except Exception:
                continue
        if ktv:
            self.KT_V_TABLE = ktv

        ktz_raw = tables.get("KT_Z_TABLE", {})
        ktz: dict[str, dict[int, float]] = {}
        for ins, inner in ktz_raw.items():
            try:
                ktz[ins] = {int(k): float(v) for k, v in inner.items()}
            except Exception:
                continue
        if ktz:
            self.KT_Z_TABLE = ktz

    def _update_medium_options(self) -> None:
        if not self._medium_combobox:
            return
        language = self._language.get()
        values = [meta.get(language, meta.get(self.DEFAULT_LANGUAGE, "")) for meta in self.TEMPERATURE_MEDIA.values()]
        self._medium_combobox.configure(values=values)
        self._combobox_values["Среда для Т"] = values
        display = self.TEMPERATURE_MEDIA[self._medium_selected_key].get(
            language, self.TEMPERATURE_MEDIA[self._medium_selected_key][self.DEFAULT_LANGUAGE]
        )
        self._medium_combobox.set(display)
        var = self._form_values.get("Среда для Т")
        if var is not None:
            var.set(display)

    def _update_help_text(self) -> None:
        if self._help_text_widget is None:
            return
        language = self._language.get()
        help_text = self.HELP_TEXTS.get(language, self.HELP_TEXTS[self.DEFAULT_LANGUAGE])
        self._help_text_widget.configure(state="normal")
        self._help_text_widget.delete("1.0", tk.END)
        self._help_text_widget.insert("1.0", help_text)
        self._help_text_widget.configure(state="disabled")

    def _on_medium_changed(self, _: tk.Event | None) -> None:
        display = self._form_values.get("Среда для Т")
        if display is None:
            return
        selected = display.get().strip()
        for key, translations in self.TEMPERATURE_MEDIA.items():
            if selected in translations.values():
                self._medium_selected_key = key
                break
        self._update_intermediate_results()

    def _set_medium_from_value(self, value: str) -> None:
        normalized = value.strip()
        for key, translations in self.TEMPERATURE_MEDIA.items():
            if normalized in translations.values():
                self._medium_selected_key = key
                break
        language = self._language.get()
        display = self.TEMPERATURE_MEDIA[self._medium_selected_key].get(
            language, self.TEMPERATURE_MEDIA[self._medium_selected_key][self.DEFAULT_LANGUAGE]
        )
        medium_var = self._form_values.get("Среда для Т")
        if medium_var is not None:
            medium_var.set(display)
        if self._medium_combobox is not None:
            self._medium_combobox.set(display)

    def _attach_tooltip(self, widget: ttk.Widget, label_key: str) -> None:
        tooltip_texts = self.TOOLTIPS.get(label_key)
        if not tooltip_texts:
            return

        def text_getter(key: str = label_key) -> str:
            translations = self.TOOLTIPS.get(key, {})
            language = self._language.get()
            if language in translations:
                return translations[language]
            return translations.get(self.DEFAULT_LANGUAGE, "")

        tooltip = Tooltip(widget, text_getter)
        self._tooltips.append(tooltip)

    def _build_menu(self) -> None:
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label=self._("menu.save_project"), command=self.save_project)
        self._register_menu_text(file_menu, file_menu.index("end"), "menu.save_project")
        file_menu.add_command(label=self._("menu.load_project"), command=self.load_project)
        self._register_menu_text(file_menu, file_menu.index("end"), "menu.load_project")
        file_menu.add_separator()
        file_menu.add_command(label=self._("menu.export_excel"), command=self.export_to_excel)
        self._register_menu_text(file_menu, file_menu.index("end"), "menu.export_excel")

        menubar.add_cascade(label=self._("menu.file"), menu=file_menu)
        self._register_menu_text(menubar, menubar.index("end"), "menu.file")

        language_menu = tk.Menu(menubar, tearoff=0)
        for code, name in self.LANGUAGES.items():
            language_menu.add_radiobutton(
                label=name,
                variable=self._language,
                value=code,
                command=self._apply_language,
            )
        menubar.add_cascade(label=self._("menu.language"), menu=language_menu)
        self._register_menu_text(menubar, menubar.index("end"), "menu.language")

        self.config(menu=menubar)

    def _build_layout(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._notebook = notebook

        main_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="")
        self._register_notebook_tab(main_tab, "tab.calculation")
        self._tabs["tab.calculation"] = main_tab

        help_tab = ttk.Frame(notebook)
        notebook.add(help_tab, text="")
        self._register_notebook_tab(help_tab, "tab.help")
        self._tabs["tab.help"] = help_tab

        container = ttk.Frame(main_tab)
        container.pack(fill=tk.BOTH, expand=True)

        form_frame = ttk.LabelFrame(container)
        form_frame.pack(fill=tk.X, expand=False, side=tk.TOP, pady=(0, 10))
        self._bind_text(lambda value, widget=form_frame: widget.configure(text=value), "frame.input")
        self._build_form(form_frame)

        intermediate_frame = ttk.LabelFrame(container)
        intermediate_frame.pack(fill=tk.X, expand=False, side=tk.TOP, pady=(0, 10))
        self._bind_text(
            lambda value, widget=intermediate_frame: widget.configure(text=value), "frame.intermediate"
        )
        self._build_intermediate_panel(intermediate_frame)

        table_frame = ttk.LabelFrame(container)
        table_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        self._bind_text(lambda value, widget=table_frame: widget.configure(text=value), "frame.table")
        self._build_table(table_frame)

        self._register_form_traces()

        self._build_help_tab(help_tab)

    def _build_help_tab(self, parent: ttk.Frame) -> None:
        text = tk.Text(parent, wrap="word", height=10)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._help_text_widget = text
        self._update_help_text()

    def _build_form(self, parent: ttk.Frame) -> None:
        default_medium_display = self.TEMPERATURE_MEDIA[self._medium_selected_key][self.DEFAULT_LANGUAGE]
        field_specs = [
            ("Strujni krug", ""),
            ("Deonica OD", ""),
            ("Deonica DO", ""),
            ("Tip-IZOLACIJE", self.INSULATION_OPTIONS[0]),
            ("Tip-PROVODNIKA", self.CONDUCTOR_TYPES[0]),
            ("Oznaka-tip-KABLA", ""),
            ("Pi, W", ""),
            ("Kj", ""),
            ("η", "1.0"),
            ("U", self.VOLTAGE_LEVELS[0]),
            ("cos φ", ""),
            ("Dužina L, m", ""),
            ("Presek, mm²", ""),
            ("Način polaganja", self.INSTALLATION_METHODS[4]),
            ("Нагруженные жилы (nž)", "3"),
            ("Кабелей в группе (для S)", "1"),
            ("Параллельные кабели (n∥)", "1"),
            ("Среда для Т", default_medium_display),
            ("Температура, °C", "30"),
            ("In, A", ""),
            ("k", "1.45"),
            ("Ключ ΔU", list(self.DROP_LIMIT_KEYS.keys())[0]),
        ]

        grid = ttk.Frame(parent)
        grid.pack(fill=tk.X, expand=False, padx=10, pady=10)

        columns = 4
        rows_per_column = math.ceil(len(field_specs) / columns)

        entry_columns = [col * 2 + 1 for col in range(columns)]
        for col in range(columns * 2):
            weight = 1 if col in entry_columns else 0
            grid.columnconfigure(col, weight=weight)

        for index, (label, default) in enumerate(field_specs):
            column = index // rows_per_column
            row = index % rows_per_column
            label_col = column * 2
            entry_col = label_col + 1

            label_widget = ttk.Label(grid, text="", anchor="e", justify="right")
            label_widget.grid(row=row, column=label_col, sticky=tk.E, pady=4, padx=(0, 8))
            label_key = self.LABEL_KEY_MAP.get(label, label)
            self._bind_text(lambda value, widget=label_widget: widget.configure(text=value), label_key)

            var = tk.StringVar(value=default)
            self._form_values[label] = var

            if label == "Tip-IZOLACIJE":
                widget = ttk.Combobox(grid, textvariable=var, values=self.INSULATION_OPTIONS, state="readonly")
            elif label == "Tip-PROVODNIKA":
                widget = ttk.Combobox(grid, textvariable=var, values=self.CONDUCTOR_TYPES, state="readonly")
            elif label == "U":
                widget = ttk.Combobox(grid, textvariable=var, values=self.VOLTAGE_LEVELS, state="readonly")
                self._combobox_values[label] = list(self.VOLTAGE_LEVELS)
            elif label == "Način polaganja":
                widget = ttk.Combobox(grid, textvariable=var, values=self.INSTALLATION_METHODS, state="readonly")
                self._combobox_values[label] = list(self.INSTALLATION_METHODS)
            elif label == "Нагруженные жилы (nž)":
                widget = ttk.Combobox(grid, textvariable=var, values=["2", "3"], state="readonly")
                self._combobox_values[label] = ["2", "3"]
            elif label == "Кабелей в группе (для S)":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=[str(i) for i in range(1, 21)], state="readonly"
                )
                self._combobox_values[label] = [str(i) for i in range(1, 21)]
            elif label == "Параллельные кабели (n∥)":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=[str(i) for i in range(1, 7)], state="readonly"
                )
                self._combobox_values[label] = [str(i) for i in range(1, 7)]
            elif label == "Среда для Т":
                medium_values = [
                    meta.get(self._language.get(), meta.get(self.DEFAULT_LANGUAGE, ""))
                    for meta in self.TEMPERATURE_MEDIA.values()
                ]
                widget = ttk.Combobox(grid, textvariable=var, values=medium_values, state="readonly")
                widget.bind("<<ComboboxSelected>>", self._on_medium_changed)
                self._medium_combobox = widget
                self._combobox_values[label] = medium_values
            elif label == "Ключ ΔU":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=list(self.DROP_LIMIT_KEYS.keys()), state="readonly"
                )
                self._combobox_values[label] = list(self.DROP_LIMIT_KEYS.keys())
            elif label == "Температура, °C":
                widget = ttk.Entry(grid, textvariable=var)
                widget.bind("<FocusIn>", self._on_temperature_focus_in)
                widget.bind("<FocusOut>", self._on_temperature_focus_out)
            elif label == "Presek, mm²":
                widget = ttk.Combobox(
                    grid,
                    textvariable=var,
                    values=self.STANDARD_CROSS_SECTIONS,
                    state="readonly",
                )
                self._combobox_values[label] = list(self.STANDARD_CROSS_SECTIONS)
            elif label == "In, A":
                widget = ttk.Combobox(
                    grid,
                    textvariable=var,
                    values=self.STANDARD_BREAKER_RATINGS,
                    state="readonly",
                )
                self._combobox_values[label] = list(self.STANDARD_BREAKER_RATINGS)
            else:
                widget = ttk.Entry(grid, textvariable=var)

            widget.grid(row=row, column=entry_col, sticky=tk.EW, pady=4)
            widget_class = widget.winfo_class()
            original_style = widget.cget("style") or widget_class
            self._input_widgets[label] = widget
            self._input_styles[label] = original_style
            self._attach_tooltip(widget, label_key)

        extra_row = rows_per_column
        check_button = ttk.Checkbutton(grid, text="", variable=self._consider_parallel_in_s)
        check_button.grid(row=extra_row, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        self._bind_text(lambda value, widget=check_button: widget.configure(text=value), "checkbox.consider_parallel")
        self._attach_tooltip(check_button, "checkbox.consider_parallel")
        self._input_widgets["Учитывать n∥ в S"] = check_button
        self._input_styles["Учитывать n∥ в S"] = check_button.cget("style") or check_button.winfo_class()

        pi_var = self._form_values["Pi, W"]
        kj_var = self._form_values["Kj"]
        pi_var.trace_add("write", self._update_pj_display)
        kj_var.trace_add("write", self._update_pj_display)

        for key, default in (("Pj", ""), ("S", "1.0"), ("T", "1.0")):
            if key not in self._form_values:
                self._form_values[key] = tk.StringVar(value=default)

    def _build_intermediate_panel(self, parent: ttk.Frame) -> None:
        grid = ttk.Frame(parent)
        grid.pack(fill=tk.X, expand=False, padx=10, pady=10)

        specs = [
            ("Pj, W", "Pj, W"),
            ("Icalc [A]", "Icalc [A]"),
            ("R_base [Ω/km]", "R_base [Ω/km]"),
            ("Iz [A]", "Iz [A]"),
            ("S", "S"),
            ("T", "T"),
            ("ΔU %", "ΔU %"),
            ("Ukupni ΔU %", "Ukupni ΔU %"),
            ("Limit ΔU %", "Limit ΔU %"),
            ("По току", "По току"),
            ("По ΔU", "По ΔU"),
            ("Диапазон In [A]", "Диапазон In [A]"),
            ("I2 [A]", "I2 [A]"),
            ("Защита", "Защита"),
            ("Совместимость IEC", "Совместимость IEC"),
            ("Рекомендации", "Рекомендации"),
        ]

        columns = 3
        for col in range(columns * 2):
            weight = 1 if col % 2 == 1 else 0
            grid.columnconfigure(col, weight=weight)

        for index, (label_text, key) in enumerate(specs):
            row = index // columns
            label_col = (index % columns) * 2
            value_col = label_col + 1

            label_widget = ttk.Label(grid, text="", style="ResultKey.TLabel")
            label_widget.grid(row=row, column=label_col, sticky=tk.W, pady=4, padx=(0, 8))
            label_key = self.RESULT_LABEL_KEY_MAP.get(label_text, label_text)
            self._bind_text(lambda value, widget=label_widget: widget.configure(text=value), label_key)

            var = tk.StringVar(value="—")
            value_label = ttk.Label(grid, textvariable=var, style="ResultValue.TLabel")
            value_label.grid(row=row, column=value_col, sticky=tk.EW, pady=4)
            self._intermediate_vars[key] = var
            self._intermediate_labels[key] = value_label

        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        select_button = ttk.Button(button_frame, text="", command=self.select_optimal_parameters)
        select_button.pack(side=tk.LEFT, padx=(0, 5))
        self._bind_text(
            lambda value, widget=select_button: widget.configure(text=value), "button.select_optimal"
        )

        add_button = ttk.Button(button_frame, text="", command=self.add_row)
        add_button.pack(side=tk.LEFT, padx=(0, 5))
        self._bind_text(lambda value, widget=add_button: widget.configure(text=value), "button.add_row")

        load_button = ttk.Button(button_frame, text="", command=self.load_selected_row)
        load_button.pack(side=tk.LEFT, padx=(0, 5))
        self._bind_text(lambda value, widget=load_button: widget.configure(text=value), "button.load_row")

        delete_button = ttk.Button(button_frame, text="", command=self.remove_selected_row)
        delete_button.pack(side=tk.LEFT, padx=(0, 5))
        self._bind_text(lambda value, widget=delete_button: widget.configure(text=value), "button.remove_row")

        clear_button = ttk.Button(button_frame, text="", command=self.clear_table)
        clear_button.pack(side=tk.LEFT)
        self._bind_text(lambda value, widget=clear_button: widget.configure(text=value), "button.clear")

    def _build_table(self, parent: ttk.Frame) -> None:
        columns = self.TREE_COLUMNS

        tree_container = ttk.Frame(parent)
        tree_container.pack(fill=tk.BOTH, expand=True)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        tree = ttk.Treeview(tree_container, columns=columns, show="headings")
        self.tree = tree

        for col in columns:
            key = self.TREE_COLUMN_KEYS.get(col, col)
            self._register_tree_heading(col, key)
            tree.column(col, width=120, anchor=tk.CENTER)

        tree.grid(row=0, column=0, sticky=tk.NSEW)

        v_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree.yview)
        v_scroll.grid(row=0, column=1, sticky=tk.NS)

        h_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=tree.xview)
        h_scroll.grid(row=1, column=0, sticky=tk.EW)

        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

    def _register_form_traces(self) -> None:
        for name, var in self._form_values.items():
            if name in {"Pj", "S", "T"}:
                continue
            var.trace_add("write", self._update_intermediate_results)
        self._consider_parallel_in_s.trace_add("write", self._update_intermediate_results)
        self._update_intermediate_results()

    def _try_parse_float(self, value: str) -> float | None:
        value = value.strip().replace(",", ".")
        if not value:
            return None
        try:
            return float(value)
        except ValueError:
            return None

    def _parse_float(self, value: str, field_name: str) -> float | None:
        value = value.strip()
        if not value:
            messagebox.showerror("Ошибка ввода", f"Поле '{field_name}' должно быть заполнено числом.")
            logging.error("Пустое значение в поле '%s'", field_name)
            return None
        try:
            return float(value.replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка ввода", f"Поле '{field_name}' содержит недопустимое значение: {value}")
            logging.error("Некорректное значение '%s' в поле '%s'", value, field_name)
            return None

    def _fmt(self, value: float | None, digits: int = 2) -> str:
        if value is None:
            return "—"
        try:
            if not math.isfinite(value):
                return "—"
        except TypeError:
            return "—"
        formatted = f"{value:.{digits}f}"
        if self._language.get() in {"ru", "sr"}:
            formatted = formatted.replace(".", ",")
        return formatted

    def _lookup_ampacity(
        self, insulation_key: str, conductor: str, laying: str, area: float, loaded_cores: int
    ) -> float | None:
        base_table = self.AMPACITY_BASE.get(laying)
        if not base_table:
            return None

        insulation_factors = self.AMPACITY_INSULATION_FACTORS.get(insulation_key, {}).get(conductor)
        load_factors = self.AMPACITY_LOADED_FACTORS.get(laying)
        if insulation_factors is None or load_factors is None:
            return None

        multiplier = insulation_factors.get(laying)
        load_multiplier = load_factors.get(loaded_cores)
        if multiplier is None or load_multiplier is None:
            return None

        standard_sections = sorted(base_table)
        if not standard_sections:
            return None

        if area <= standard_sections[0]:
            base_value = base_table[standard_sections[0]]
        elif area >= standard_sections[-1]:
            base_value = base_table[standard_sections[-1]]
        else:
            for lower, upper in zip(standard_sections, standard_sections[1:]):
                if math.isclose(area, lower, rel_tol=1e-6, abs_tol=1e-3):
                    base_value = base_table[lower]
                    break
                if lower <= area <= upper:
                    lower_val = base_table[lower]
                    upper_val = base_table[upper]
                    if math.isclose(upper, lower):
                        base_value = lower_val
                    else:
                        ratio = (area - lower) / (upper - lower)
                        base_value = lower_val + ratio * (upper_val - lower_val)
                    break
            else:
                return None

        return base_value * multiplier * load_multiplier

    def _lookup_group_factor(self, circuits: int) -> float:
        if circuits <= 1:
            return 1.0
        max_defined = max(self.GROUPING_FACTORS)
        if circuits >= max_defined:
            return self.GROUPING_FACTORS[max_defined]
        return self.GROUPING_FACTORS.get(circuits, self.GROUPING_FACTORS[max_defined])

    def _lookup_temperature_factor(self, insulation_key: str, medium: str, temperature: float) -> float | None:
        if medium == "soil":
            table = self.KT_Z_TABLE.get(insulation_key)
        else:
            table = self.KT_V_TABLE.get(insulation_key)
        if not table:
            return None

        points = sorted(table)
        if not points:
            return None
        if temperature < points[0] or temperature > points[-1]:
            return None

        if temperature in table:
            return table[temperature]

        for lower, upper in zip(points, points[1:]):
            if lower <= temperature <= upper:
                lower_val = table[lower]
                upper_val = table[upper]
                if math.isclose(upper, lower):
                    return lower_val
                ratio = (temperature - lower) / (upper - lower)
                return lower_val + ratio * (upper_val - lower_val)
        return None

    def _calculate_line_impedance(
        self, conductor: str, insulation_temp: float, area: float, laying: str
    ) -> tuple[float, float]:
        rho_20 = self.RESISTIVITY_20.get(conductor)
        alpha = self.TEMP_COEFF.get(conductor)
        if rho_20 is None or alpha is None:
            default_x = 0.08
            if isinstance(getattr(self, "REACTANCE_DATA", None), dict) and self.REACTANCE_DATA:
                default_x = self.REACTANCE_DATA.get("default", default_x)
            return 0.0, default_x

        rho_theta = rho_20 * (1.0 + alpha * (insulation_temp - 20.0))
        if area <= 0:
            r_per_km = 0.0
        else:
            r_per_km = (rho_theta / area) * 1000.0

        bucket = "≤95"
        if area > 240:
            bucket = ">240"
        elif area > 95:
            bucket = "≤240"
        # Prefer externally loaded reactance data if available
        x_per_km = None
        if isinstance(getattr(self, "REACTANCE_DATA", None), dict) and self.REACTANCE_DATA:
            try:
                buckets = self.REACTANCE_DATA.get("buckets", {})
                method_defaults = self.REACTANCE_DATA.get("method_defaults", {})
                x_per_km = buckets.get(laying, {}).get(bucket)
                if x_per_km is None:
                    x_per_km = method_defaults.get(laying)
                if x_per_km is None:
                    x_per_km = self.REACTANCE_DATA.get("default")
            except Exception:
                x_per_km = None
        if x_per_km is None:
            x_per_km = self.REACTANCE_PER_KM.get((laying, bucket))
            if x_per_km is None:
                x_per_km = self.REACTANCE_PER_KM.get(laying)
            if x_per_km is None:
                x_per_km = self.REACTANCE_PER_KM.get("default", 0.08)
        return r_per_km, x_per_km

    def _drop_pct(
        self,
        U: float,
        cos_phi: float,
        L_m: float,
        conductor: str,
        insulation_theta: float,
        area_mm2: float,
        method: str,
        loaded_cores: int,
        n_parallel: int,
        icalc_total: float,
    ) -> float:
        r_km, x_km = self._calculate_line_impedance(conductor, insulation_theta, area_mm2, method)
        divider = max(n_parallel, 1)
        r_m, x_m = (r_km / 1000.0) / divider, (x_km / 1000.0) / divider
        sin_phi = math.sqrt(max(0.0, 1.0 - min(1.0, cos_phi) ** 2))
        phase = 2.0 if loaded_cores == 2 else math.sqrt(3)
        if U == 0:
            return 0.0
        return phase * icalc_total * (r_m * cos_phi + x_m * sin_phi) * L_m * 100.0 / U

    def _recommend(
        self,
        *,
        U: float,
        cos_phi: float,
        L: float,
        conductor: str,
        insulation_key: str,
        insulation_theta: float,
        method: str,
        loaded_cores: int,
        S: float,
        T: float,
        limit_pct: float | None,
        current_area: float,
        n_parallel: int,
        icalc_total: float,
    ) -> list[str]:
        recs: list[str] = []
        if not icalc_total or S <= 0 or T <= 0:
            return recs

        parallel_count = max(n_parallel, 1)
        current_per_cable = icalc_total / parallel_count if parallel_count else icalc_total
        min_section = max(current_area or 0, self.STANDARD_SECTIONS[0])

        def fmt(value: float, digits: int = 2) -> str:
            return self._fmt(value, digits=digits)

        def candidate_description(action: str, area_val: float, method_val: str, iz_one: float, drop_val: float) -> str:
            digits = 0 if float(area_val).is_integer() else 1
            area_str = fmt(area_val, digits=digits)
            iz_total = iz_one * parallel_count
            return (
                f"{action} {area_str} мм² (метод {method_val}) → "
                f"Iz_tot≈{fmt(iz_total, digits=0)} A, ΔU≈{fmt(drop_val)}%"
            )

        # 1) Increase section within same method/insulation
        best_section_line = None
        for area_candidate in self.STANDARD_SECTIONS:
            if area_candidate < min_section:
                continue
            iz_base = self._lookup_ampacity(insulation_key, conductor, method, area_candidate, loaded_cores)
            if not iz_base:
                continue
            iz_one = iz_base * S * T
            if iz_one < current_per_cable:
                continue
            drop_val = self._drop_pct(
                U,
                cos_phi,
                L,
                conductor,
                insulation_theta,
                area_candidate,
                method,
                loaded_cores,
                parallel_count,
                icalc_total,
            )
            if limit_pct is not None and drop_val > limit_pct:
                continue
            best_section_line = candidate_description("Увеличить сечение до", area_candidate, method, iz_one, drop_val)
            break
        if best_section_line:
            recs.append(best_section_line)

        # 2) Change method according to preference
        best_method_line = None
        for m in self.METHOD_PREFERENCE:
            if m == method:
                continue
            for area_candidate in self.STANDARD_SECTIONS:
                if area_candidate < min_section:
                    continue
                iz_base = self._lookup_ampacity(insulation_key, conductor, m, area_candidate, loaded_cores)
                if not iz_base:
                    continue
                iz_one = iz_base * S * T
                if iz_one < current_per_cable:
                    continue
                drop_val = self._drop_pct(
                    U,
                    cos_phi,
                    L,
                    conductor,
                    insulation_theta,
                    area_candidate,
                    m,
                    loaded_cores,
                    parallel_count,
                    icalc_total,
                )
                if limit_pct is not None and drop_val > limit_pct:
                    continue
                best_method_line = candidate_description("Сменить метод на", area_candidate, m, iz_one, drop_val)
                break
            if best_method_line:
                break
        if best_method_line:
            recs.append(best_method_line)

        # 3) Switch to XLPE if currently PVC
        if insulation_key == "PVC":
            xlpe_meta = self.INSULATION_META.get("XLPE/EPR (90°C)")
            xlpe_theta = xlpe_meta.get("theta", 90.0) if xlpe_meta else 90.0
            best_xlpe_line = None
            for area_candidate in self.STANDARD_SECTIONS:
                if area_candidate < min_section:
                    continue
                iz_base = self._lookup_ampacity("XLPE", conductor, method, area_candidate, loaded_cores)
                if not iz_base:
                    continue
                iz_one = iz_base * S * T
                if iz_one < current_per_cable:
                    continue
                drop_val = self._drop_pct(
                    U,
                    cos_phi,
                    L,
                    conductor,
                    xlpe_theta,
                    area_candidate,
                    method,
                    loaded_cores,
                    parallel_count,
                    icalc_total,
                )
                if limit_pct is not None and drop_val > limit_pct:
                    continue
                digits = 0 if float(area_candidate).is_integer() else 1
                area_str = fmt(area_candidate, digits=digits)
                iz_total = iz_one * parallel_count
                best_xlpe_line = (
                    f"Перейти на XLPE и {area_str} мм² (метод {method}) → "
                    f"Iz_tot≈{fmt(iz_total, digits=0)} A, ΔU≈{fmt(drop_val)}%"
                )
                break
            if best_xlpe_line:
                recs.append(best_xlpe_line)

        # 4) Increase number of parallel cables to meet voltage drop
        if limit_pct is not None:
            base_area = max(min_section, current_area or self.STANDARD_SECTIONS[0])
            drop_current = self._drop_pct(
                U,
                cos_phi,
                L,
                conductor,
                insulation_theta,
                base_area,
                method,
                loaded_cores,
                parallel_count,
                icalc_total,
            )
            if drop_current > limit_pct:
                iz_base = self._lookup_ampacity(insulation_key, conductor, method, base_area, loaded_cores)
                if iz_base:
                    iz_one = iz_base * S * T
                    denom = max(limit_pct, 1e-9)
                    n_needed = max(parallel_count + 1, math.ceil(drop_current * parallel_count / denom))
                    if iz_one > 0:
                        max_total = iz_one * n_needed
                        drop_new = self._drop_pct(
                            U,
                            cos_phi,
                            L,
                            conductor,
                            insulation_theta,
                            base_area,
                            method,
                            loaded_cores,
                            n_needed,
                            icalc_total,
                        )
                        digits = 0 if float(base_area).is_integer() else 1
                        area_str = fmt(base_area, digits=digits)
                        recs.append(
                            f"Разделить на {n_needed} параллельных кабеля {area_str} мм² (метод {method}) → "
                            f"n∥={n_needed}, Iz_tot≈{fmt(max_total, digits=0)} A, ΔU≈{fmt(drop_new)}%"
                        )

        if not recs:
            recs.append("Снизить число кабелей в группе (для увеличения S) или повысить напряжение.")
        return recs[:4]

    def _update_pj_display(self, *_: object) -> None:
        pi = self._try_parse_float(self._form_values["Pi, W"].get())
        kj = self._try_parse_float(self._form_values["Kj"].get())
        if pi is None or kj is None:
            self._form_values["Pj"].set("")
            self._update_intermediate_results()
            return
        self._form_values["Pj"].set(self._fmt(pi * kj))
        self._update_intermediate_results()

    def _set_result_alert(self, key: str, alert: bool) -> None:
        label = self._intermediate_labels.get(key)
        if not label:
            return
        label.configure(style="ResultAlert.TLabel" if alert else "ResultValue.TLabel")

    def _set_entry_alert(self, field_name: str, alert: bool) -> None:
        widget = self._input_widgets.get(field_name)
        if widget is None:
            return
        default_style = self._input_styles.get(field_name, "")
        if not alert:
            widget.configure(style=default_style)
            return
        widget_class = widget.winfo_class()
        if widget_class == "TEntry":
            widget.configure(style="Invalid.TEntry")
        elif widget_class == "TCombobox":
            widget.configure(style="Invalid.TCombobox")

    def _on_temperature_focus_in(self, _: tk.Event) -> None:
        self._temperature_editing = True

    def _on_temperature_focus_out(self, _: tk.Event) -> None:
        self._temperature_editing = False
        self._update_intermediate_results()

    def _show_temperature_warning(self, insulation_key: str, medium: str, temperature: float) -> None:
        if self._temperature_editing:
            widget = self._input_widgets.get("Температура, °C")
            if widget is not None and self.focus_get() == widget:
                return
        rounded_temp = round(temperature, 1)
        key = (insulation_key, medium, rounded_temp)
        if self._last_temperature_warning == key:
            return
        self._last_temperature_warning = key
        logging.warning(
            "Температура вне диапазона для изоляции %s, среды %s: %s °C",
            insulation_key,
            medium,
            rounded_temp,
        )
        messagebox.showwarning(
            "Температура вне диапазона",
            "Для выбранной изоляции и среды отсутствует табличный коэффициент при температуре "
            f"{rounded_temp} °C. Проверьте корректность условий или используйте значения в пределах таблиц IEC 60364.",
        )

    def _update_intermediate_results(self, *_: object) -> None:
        if not self._intermediate_vars:
            return

        self._last_icalc = None
        for key, var in self._intermediate_vars.items():
            if key == "Limit ΔU %":
                continue
            var.set("—")

        for key in ("ΔU %", "По ΔU", "По току", "Защита", "Ukupni ΔU %"):
            self._set_result_alert(key, False)

        strujni_krug = self._form_values["Strujni krug"].get().strip()
        insulation_label = self._form_values["Tip-IZOLACIJE"].get()
        insulation_meta = self.INSULATION_META.get(insulation_label)
        conductor = self._form_values["Tip-PROVODNIKA"].get()
        laying = self._form_values["Način polaganja"].get().strip()
        loaded_cores_value = self._form_values["Нагруженные жилы (nž)"].get().strip()
        group_value = self._form_values["Кабелей в группе (для S)"].get().strip()
        parallel_value = self._form_values["Параллельные кабели (n∥)"].get().strip()
        if self._medium_selected_key not in self.TEMPERATURE_MEDIA:
            self._medium_selected_key = next(iter(self.TEMPERATURE_MEDIA))
            self._update_medium_options()
        medium_key = self._medium_selected_key
        drop_key = self._form_values["Ключ ΔU"].get()
        limit_delta = self.DROP_LIMIT_KEYS.get(drop_key)
        if limit_delta is not None:
            self._intermediate_vars["Limit ΔU %"].set(self._fmt(limit_delta))
        else:
            self._intermediate_vars["Limit ΔU %"].set("—")

        loaded_cores = 3
        cores_alert = False
        try:
            loaded_cores = int(loaded_cores_value)
            if loaded_cores not in (2, 3):
                raise ValueError
        except ValueError:
            cores_alert = True
            loaded_cores = 3
        self._set_entry_alert("Нагруженные жилы (nž)", cores_alert)

        circuits_alert = False
        try:
            circuits_count = int(group_value)
            if circuits_count < 1:
                raise ValueError
        except ValueError:
            circuits_count = 1
            circuits_alert = True
        self._set_entry_alert("Кабелей в группе (для S)", circuits_alert)

        parallel_alert = False
        try:
            n_parallel = int(parallel_value)
            if n_parallel < 1:
                raise ValueError
        except ValueError:
            n_parallel = 1
            parallel_alert = True
        self._set_entry_alert("Параллельные кабели (n∥)", parallel_alert)

        pi = self._try_parse_float(self._form_values["Pi, W"].get())
        kj = self._try_parse_float(self._form_values["Kj"].get())
        eta_value = self._try_parse_float(self._form_values["η"].get())
        cos_phi = self._try_parse_float(self._form_values["cos φ"].get())
        length = self._try_parse_float(self._form_values["Dužina L, m"].get())
        area = self._try_parse_float(self._form_values["Presek, mm²"].get())
        temperature = self._try_parse_float(self._form_values["Температура, °C"].get())

        effective_circuits = circuits_count
        if self._consider_parallel_in_s.get():
            effective_circuits = max(1, circuits_count + n_parallel - 1)
        group_factor = self._lookup_group_factor(effective_circuits)
        s_display = self._fmt(group_factor)
        self._form_values["S"].set(s_display if s_display != "—" else "")
        self._intermediate_vars["S"].set(s_display)
        s_coeff = group_factor

        temperature_alert = False
        t_coeff: float | None = None
        if insulation_meta is not None and temperature is not None:
            temp_factor = self._lookup_temperature_factor(insulation_meta["key"], medium_key, temperature)
            if temp_factor is not None and temp_factor > 0:
                t_coeff = temp_factor
            else:
                temperature_alert = True
        elif insulation_meta is not None:
            temperature_alert = True

        if t_coeff is not None:
            t_display = self._fmt(t_coeff)
            self._last_temperature_warning = None
        else:
            t_display = ""
            if (
                temperature is not None
                and insulation_meta is not None
                and not self._temperature_editing
            ):
                self._show_temperature_warning(insulation_meta["key"], medium_key, temperature)
        if t_display:
            self._form_values["T"].set(t_display)
            self._intermediate_vars["T"].set(t_display)
        else:
            self._form_values["T"].set("")
            self._intermediate_vars["T"].set("—")
        if self._temperature_editing:
            temperature_alert = False
        self._set_entry_alert("Температура, °C", temperature_alert)

        eta_alert = False
        eta_coeff: float | None = None
        if eta_value is not None:
            if eta_value <= 0 or eta_value > 1:
                eta_alert = True
            else:
                eta_coeff = eta_value
        else:
            eta_alert = True
        self._set_entry_alert("η", eta_alert)

        voltage_str = self._form_values["U"].get()
        try:
            voltage_value = int(voltage_str)
        except (TypeError, ValueError):
            voltage_value = None

        voltage_alert = voltage_value == 230 and loaded_cores == 3
        if voltage_alert and not self._voltage_phase_warning_shown:
            messagebox.showwarning("IEC 60364", self._("warning.voltage_phase"))
            self._voltage_phase_warning_shown = True
        self._set_entry_alert("U", voltage_alert)

        pj = None
        if pi is not None and kj is not None:
            pj = pi * kj
            self._intermediate_vars["Pj, W"].set(self._fmt(pj))

        cos_alert = False
        if cos_phi is not None:
            if not (0 < cos_phi <= 1):
                cos_alert = True
                cos_phi = None
        self._set_entry_alert("cos φ", cos_alert)

        area_alert = False
        if area is not None and area <= 0:
            area_alert = True
            area = None

        length_alert = False
        if length is not None and length < 0:
            length_alert = True
            length = None

        self._set_entry_alert("Presek, mm²", area_alert)
        self._set_entry_alert("Dužina L, m", length_alert)

        in_range_alert = False
        icalc_total = None
        icalc_per_cable = None
        phase_factor = None
        if pj is not None and cos_phi is not None and voltage_value and eta_coeff is not None:
            phase_factor = 2.0 if loaded_cores == 2 else math.sqrt(3)
            denominator = phase_factor * voltage_value * cos_phi
            if denominator:
                icalc_total = (pj / eta_coeff) / denominator
                self._intermediate_vars["Icalc [A]"].set(self._fmt(icalc_total, digits=3))
                self._last_icalc = icalc_total
                icalc_per_cable = icalc_total / max(n_parallel, 1)
        else:
            self._intermediate_vars["Icalc [A]"].set("—")

        r_per_km = None
        x_per_km = None
        if area is not None and insulation_meta is not None:
            r_per_km, x_per_km = self._calculate_line_impedance(
                conductor, insulation_meta["theta"], area, laying
            )
            self._intermediate_vars["R_base [Ω/km]"].set(self._fmt(r_per_km, digits=3))

        base_ampacity = None
        if area is not None and insulation_meta is not None:
            base_ampacity = self._lookup_ampacity(
                insulation_meta["key"], conductor, laying, area, loaded_cores
            )

        iz_one = None
        if base_ampacity is not None and t_coeff is not None:
            iz_one = base_ampacity * s_coeff * t_coeff
            self._intermediate_vars["Iz [A]"].set(self._fmt(iz_one))
        elif base_ampacity is None:
            self._intermediate_vars["Iz [A]"].set("—")

        in_range_value = "—"
        if icalc_total is not None and iz_one is not None:
            iz_total = iz_one * n_parallel
            in_range_value = f"{self._fmt(icalc_total)} – {self._fmt(iz_total)}"
            if iz_total + 1e-9 < icalc_total:
                in_range_alert = True
        if "Диапазон In [A]" in self._intermediate_vars:
            self._intermediate_vars["Диапазон In [A]"].set(in_range_value)

        ampacity_status = None
        if base_ampacity is None:
            ampacity_status = "N/A"
        elif iz_one is None or icalc_per_cable is None:
            ampacity_status = "—"
        else:
            ampacity_status = "OK" if icalc_per_cable <= iz_one else "NE"
        ampacity_alert = ampacity_status == "NE" or area_alert
        if ampacity_status is not None:
            self._intermediate_vars["По току"].set(ampacity_status)
            self._set_result_alert("По току", ampacity_alert)
        self._set_entry_alert("Presek, mm²", ampacity_alert)

        delta_u = None
        if (
            icalc_total is not None
            and phase_factor is not None
            and length is not None
            and cos_phi is not None
            and r_per_km is not None
            and x_per_km is not None
            and voltage_value
        ):
            divider = max(n_parallel, 1)
            r_per_meter = (r_per_km / 1000.0) / divider
            x_per_meter = (x_per_km / 1000.0) / divider
            sin_phi = math.sqrt(max(0.0, 1.0 - min(1.0, cos_phi) ** 2))
            impedance_drop = r_per_meter * cos_phi + x_per_meter * sin_phi
            delta_u = phase_factor * icalc_total * impedance_drop * length * 100.0 / voltage_value
            self._intermediate_vars["ΔU %"].set(self._fmt(delta_u))
        elif "ΔU %" in self._intermediate_vars:
            self._intermediate_vars["ΔU %"].set("—")

        drop_status = None
        if delta_u is not None and limit_delta is not None:
            drop_status = "OK" if delta_u <= limit_delta else "NE"
            self._intermediate_vars["По ΔU"].set(drop_status)
            self._set_result_alert("По ΔU", drop_status == "NE")
            self._set_result_alert("ΔU %", drop_status == "NE")
        elif delta_u is not None:
            drop_status = "—"
            self._intermediate_vars["По ΔU"].set(drop_status)

        drop_alert = (drop_status == "NE") or length_alert
        self._set_entry_alert("Dužina L, m", drop_alert)

        existing_drop = self._sum_drop_for_circuit(strujni_krug)
        if delta_u is not None:
            total_drop = existing_drop + delta_u
            self._intermediate_vars["Ukupni ΔU %"].set(self._fmt(total_drop))
            if limit_delta is not None:
                total_status = "OK" if total_drop <= limit_delta else "NE"
                self._set_result_alert("Ukupni ΔU %", total_status == "NE")
            else:
                self._set_result_alert("Ukupni ΔU %", False)
        elif existing_drop > 0:
            self._intermediate_vars["Ukupni ΔU %"].set(self._fmt(existing_drop))

        in_value = self._try_parse_float(self._form_values["In, A"].get())
        k_value = self._try_parse_float(self._form_values["k"].get())
        i2_value = None
        protection_status = "—"
        protection_alert = False

        if in_value is not None and in_value <= 0:
            protection_alert = True
            in_value = None

        if k_value is not None and k_value <= 0:
            protection_alert = True
            k_value = None

        if in_value is not None and k_value is not None:
            i2_value = in_value * k_value
            self._intermediate_vars["I2 [A]"].set(self._fmt(i2_value))
        else:
            self._intermediate_vars["I2 [A]"].set("—")

        if iz_one is None or icalc_total is None or in_value is None or i2_value is None:
            if in_value is None or i2_value is None:
                protection_status = "—"
            else:
                protection_status = "N/A"
        else:
            iz_total = iz_one * n_parallel
            within_nominal = icalc_total <= in_value <= iz_total
            overload_check = i2_value <= 1.45 * iz_total
            if within_nominal and overload_check:
                protection_status = "OK"
            else:
                protection_status = "NE"
                protection_alert = True

        self._intermediate_vars["Защита"].set(protection_status)
        self._set_result_alert("Защита", protection_alert)
        self._set_result_alert("Диапазон In [A]", in_range_alert or protection_alert)
        self._set_entry_alert("In, A", protection_alert)

        if "Совместимость IEC" in self._intermediate_vars:
            if base_ampacity is None:
                compat_display = self._("status.na")
                compat_alert = False
            else:
                statuses = [ampacity_status, drop_status, protection_status]
                if all(status == "OK" for status in statuses):
                    compat_display = self._("status.ok")
                    compat_alert = False
                elif any(status == "NE" for status in statuses):
                    compat_display = self._("status.fail")
                    compat_alert = True
                else:
                    compat_display = self._("status.na")
                    compat_alert = False
            self._intermediate_vars["Совместимость IEC"].set(compat_display)
            self._set_result_alert("Совместимость IEC", compat_alert)

        recommendations: list[str] = []
        rec_var = self._intermediate_vars.get("Рекомендации")
        if rec_var is not None:
            ampacity_state = self._intermediate_vars.get("По току")
            drop_state = self._intermediate_vars.get("По ΔU")
            ampacity_value = ampacity_state.get() if ampacity_state is not None else ""
            drop_value = drop_state.get() if drop_state is not None else ""
            if ampacity_value == "NE" or drop_value == "NE":
                insulation_label = self._form_values["Tip-IZOLACIJE"].get()
                insulation_meta = self.INSULATION_META.get(
                    insulation_label, {"key": "PVC", "theta": 70}
                )
                try:
                    loaded_cores_value = int(
                        self._form_values["Нагруженные жилы (nž)"].get() or "3"
                    )
                except (TypeError, ValueError):
                    loaded_cores_value = 3
                s_value = self._try_parse_float(self._form_values["S"].get())
                t_value = self._try_parse_float(self._form_values["T"].get())
                recs = self._recommend(
                    U=float(voltage_value or 0),
                    cos_phi=float(cos_phi or 0),
                    L=float(length or 0),
                    conductor=self._form_values["Tip-PROVODNIKA"].get(),
                    insulation_key=insulation_meta.get("key", "PVC"),
                    insulation_theta=float(insulation_meta.get("theta", 70)),
                    method=self._form_values["Način polaganja"].get().strip(),
                    loaded_cores=loaded_cores_value,
                    S=s_value if s_value is not None else 1.0,
                    T=t_value if t_value is not None else 1.0,
                    limit_pct=self.DROP_LIMIT_KEYS.get(self._form_values["Ключ ΔU"].get()),
                    current_area=self._try_parse_float(self._form_values["Presek, mm²"].get())
                    or 0,
                    n_parallel=n_parallel,
                    icalc_total=self._last_icalc or 0,
                )
                recommendations = recs or []
            rec_var.set("\n".join(recommendations) if recommendations else "—")

    def _sum_drop_for_circuit(self, circuit: str) -> float:
        if not circuit:
            return 0.0
        total = 0.0
        for row in self._table_data:
            if row.get("Strujni krug", "").strip() == circuit:
                try:
                    total += float(str(row.get("ΔU %", "0")).replace(",", "."))
                except (TypeError, ValueError):
                    continue
        return total

    def select_optimal_parameters(self) -> None:
        self._update_intermediate_results()

        insulation_label = self._form_values["Tip-IZOLACIJE"].get()
        insulation_meta = self.INSULATION_META.get(insulation_label)
        if not insulation_meta:
            messagebox.showerror("IEC 60364", "Не выбрана изоляция кабеля.")
            return

        conductor = self._form_values["Tip-PROVODNIKA"].get()
        method = self._form_values["Način polaganja"].get().strip()

        try:
            loaded_cores = int(self._form_values["Нагруженные жилы (nž)"].get())
            if loaded_cores not in (2, 3):
                raise ValueError
        except ValueError:
            messagebox.showerror("IEC 60364", "Укажите число нагруженных жил 2 или 3.")
            return

        try:
            circuits_count = int(self._form_values["Кабелей в группе (для S)"].get())
            if circuits_count < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("IEC 60364", "Число кабелей в группе должно быть не менее 1.")
            return

        try:
            n_parallel = int(self._form_values["Параллельные кабели (n∥)"].get())
            if n_parallel < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("IEC 60364", "Число параллельных кабелей должно быть не менее 1.")
            return

        pi = self._parse_float(self._form_values["Pi, W"].get(), "Pi, W")
        if pi is None:
            return
        kj = self._parse_float(self._form_values["Kj"].get(), "Kj")
        if kj is None:
            return
        eta = self._parse_float(self._form_values["η"].get(), "η")
        if eta is None:
            return
        if not (0 < eta <= 1):
            messagebox.showerror("IEC 60364", "КПД η должен быть в диапазоне (0;1].")
            return

        cos_phi = self._parse_float(self._form_values["cos φ"].get(), "cos φ")
        if cos_phi is None or not (0 < cos_phi <= 1):
            messagebox.showerror("IEC 60364", "cos φ должен быть в диапазоне (0;1].")
            return

        length = self._parse_float(self._form_values["Dužina L, m"].get(), "Dužina L, m")
        if length is None or length < 0:
            messagebox.showerror("IEC 60364", "Длина линии должна быть неотрицательной.")
            return

        voltage_str = self._form_values["U"].get()
        try:
            voltage_value = int(voltage_str)
        except (TypeError, ValueError):
            messagebox.showerror("IEC 60364", "Выберите номинальное напряжение.")
            return

        medium_key = self._medium_selected_key
        temperature = self._try_parse_float(self._form_values["Температура, °C"].get())
        t_coeff = 1.0
        if temperature is not None:
            temp_factor = self._lookup_temperature_factor(insulation_meta["key"], medium_key, temperature)
            if temp_factor is None or temp_factor <= 0:
                messagebox.showerror(
                    "IEC 60364",
                    "Температура вне табличного диапазона IEC 60364. Уточните значение, чтобы вычислить коэффициент Kt.",
                )
                return
            t_coeff = temp_factor

        effective_circuits = circuits_count
        if self._consider_parallel_in_s.get():
            effective_circuits = max(1, circuits_count + n_parallel - 1)
        s_coeff = self._lookup_group_factor(effective_circuits)

        pj = pi * kj
        phase_factor = 2.0 if loaded_cores == 2 else math.sqrt(3)
        denominator = phase_factor * voltage_value * cos_phi
        if not denominator:
            messagebox.showerror("IEC 60364", "Комбинация параметров приводит к делению на ноль.")
            return
        icalc_total = (pj / eta) / denominator

        if icalc_total <= 0:
            messagebox.showerror("IEC 60364", "Расчётный ток должен быть больше нуля для подбора параметров.")
            return

        limit_pct = self.DROP_LIMIT_KEYS.get(self._form_values["Ключ ΔU"].get())
        k_value = self._try_parse_float(self._form_values["k"].get())
        if k_value is None or k_value <= 0:
            k_value = 1.45

        breaker_values: list[tuple[str, float]] = []
        for item in self.STANDARD_BREAKER_RATINGS:
            item = item.strip()
            if not item:
                continue
            try:
                breaker_values.append((item, float(item.replace(",", "."))))
            except ValueError:
                continue

        if not breaker_values:
            messagebox.showerror("IEC 60364", "Не заданы номиналы автоматов защиты.")
            return

        insulation_key = insulation_meta["key"]
        insulation_theta = insulation_meta["theta"]

        success_combo: tuple[float, str, float, float] | None = None
        fallback_candidates: list[tuple[float, float, float, float, float, str]] = []

        for area_candidate in self.STANDARD_SECTIONS:
            if area_candidate <= 0:
                continue
            iz_base = self._lookup_ampacity(insulation_key, conductor, method, area_candidate, loaded_cores)
            if not iz_base:
                continue
            iz_one = iz_base * s_coeff * t_coeff
            iz_total = iz_one * max(n_parallel, 1)
            if iz_total <= 0:
                continue
            drop_val = self._drop_pct(
                voltage_value,
                cos_phi,
                length,
                conductor,
                insulation_theta,
                area_candidate,
                method,
                loaded_cores,
                n_parallel,
                icalc_total,
            )
            for breaker_str, breaker_value in breaker_values:
                if breaker_value <= 0:
                    continue
                within_current = icalc_total <= breaker_value <= iz_total
                drop_ok = limit_pct is None or drop_val <= limit_pct
                i2_value = breaker_value * k_value
                protection_ok = i2_value <= 1.45 * iz_total
                if within_current and drop_ok and protection_ok:
                    success_combo = (area_candidate, breaker_str, iz_total, drop_val)
                    break

                over_in_low = max(0.0, icalc_total - breaker_value)
                over_in_high = max(0.0, breaker_value - iz_total)
                over_iz = max(0.0, icalc_total - iz_total)
                over_drop = max(0.0, drop_val - (limit_pct or drop_val)) if limit_pct is not None else 0.0
                over_i2 = max(0.0, i2_value - 1.45 * iz_total)
                metric = over_in_low + over_in_high + over_iz + over_drop + over_i2
                fallback_candidates.append(
                    (
                        metric,
                        area_candidate,
                        breaker_value,
                        drop_val,
                        iz_total,
                        breaker_str,
                    )
                )
            if success_combo:
                break

        if success_combo:
            area_candidate, breaker_str, iz_total, drop_val = success_combo
            if float(area_candidate).is_integer():
                area_text = str(int(area_candidate))
            else:
                area_text = str(area_candidate)
            self._form_values["Presek, mm²"].set(area_text)
            self._form_values["In, A"].set(breaker_str)
            self._update_intermediate_results()
            return

        if not fallback_candidates:
            messagebox.showinfo("IEC 60364", self._("message.select_fail") + "нет доступных комбинаций.")
            return

        fallback_candidates.sort(key=lambda item: (item[0], item[1], item[2]))
        top_items = fallback_candidates[:3]
        lines: list[str] = []
        for metric, area_candidate, breaker_value, drop_val, iz_total, breaker_str in top_items:
            digits = 0 if float(area_candidate).is_integer() else 1
            area_display = self._fmt(area_candidate, digits=digits)
            iz_display = self._fmt(iz_total, digits=0)
            drop_display = self._fmt(drop_val)
            issues: list[str] = []
            if limit_pct is not None and drop_val > limit_pct:
                issues.append(f"ΔU +{self._fmt(drop_val - limit_pct)}%")
            if breaker_value < icalc_total:
                issues.append(f"In < Ib на {self._fmt(icalc_total - breaker_value)} A")
            if breaker_value > iz_total:
                issues.append(f"In > Iz_tot на {self._fmt(breaker_value - iz_total)} A")
            if breaker_value * k_value > 1.45 * iz_total:
                issues.append(f"I2>{self._fmt(1.45 * iz_total)} A")
            issue_text = "; ".join(issues) if issues else "минимальные отклонения"
            lines.append(
                f"• {area_display} мм² / In={breaker_str} A → Iz_tot≈{iz_display} A, ΔU≈{drop_display}% ({issue_text})"
            )

        message = self._("message.select_fail") + "\n".join(lines)
        messagebox.showinfo("IEC 60364", message)

    def remove_selected_row(self) -> None:
        if not hasattr(self, "tree"):
            return
        selected = self.tree.selection()
        if not selected:
            return
        indexed = sorted((self.tree.index(item), item) for item in selected)
        for _, item in indexed:
            self.tree.delete(item)
        for index, _ in reversed(indexed):
            if 0 <= index < len(self._table_data):
                del self._table_data[index]
        self._update_intermediate_results()

    def load_selected_row(self) -> None:
        if not hasattr(self, "tree"):
            return
        selected = self.tree.selection()
        if not selected:
            return
        row_index = self.tree.index(selected[0])
        if not (0 <= row_index < len(self._table_data)):
            return
        row = self._table_data[row_index]

        field_map = {
            "Strujni krug": "Strujni krug",
            "Deonica OD": "OD",
            "Deonica DO": "DO",
            "Tip-IZOLACIJE": "E",
            "Tip-PROVODNIKA": "F",
            "Oznaka-tip-KABLA": "G",
            "Pi, W": "Pi",
            "Kj": "Kj",
            "η": "η",
            "U": "U",
            "cos φ": "cosφ",
            "Dužina L, m": "L",
            "Presek, mm²": "Presek",
            "Način polaganja": "Način polaganja",
            "Нагруженные жилы (nž)": "nž",
            "Кабелей в группе (для S)": "Кабелей в группе (S)",
            "Параллельные кабели (n∥)": "n∥",
            "In, A": "In [A]",
            "k": "k",
            "Ключ ΔU": "Ключ",
        }

        current_insulation = self._form_values.get("Tip-IZOLACIJE")
        current_insulation_value = current_insulation.get() if current_insulation else ""

        for field, column in field_map.items():
            if field not in self._form_values:
                continue
            value = str(row.get(column, ""))
            if field == "Presek, mm²" and value:
                numeric = self._try_parse_float(value)
                if numeric is not None:
                    value = str(int(numeric)) if float(numeric).is_integer() else str(numeric)
            if field in {"Pi, W", "Kj", "η", "cos φ", "Dužina L, m"} and value:
                numeric = self._try_parse_float(value)
                if numeric is not None:
                    value = str(numeric)
            widget = self._input_widgets.get(field)
            if isinstance(widget, ttk.Combobox):
                current_values = list(widget.cget("values"))
                if value and value not in current_values:
                    current_values.append(value)
                    widget.configure(values=current_values)
                    self._combobox_values[field] = current_values
            if field == "Tip-IZOLACIJE" and value:
                if value in self.INSULATION_META:
                    self._form_values[field].set(value)
                else:
                    self._form_values[field].set(current_insulation_value)
            else:
                self._form_values[field].set(value)

        self._update_intermediate_results()

    def add_row(self) -> None:
        self._update_intermediate_results()
        strujni_krug = self._form_values["Strujni krug"].get().strip()
        od = self._form_values["Deonica OD"].get().strip()
        do = self._form_values["Deonica DO"].get().strip()
        insulation_label = self._form_values["Tip-IZOLACIJE"].get()
        conductor = self._form_values["Tip-PROVODNIKA"].get()
        cable = self._form_values["Oznaka-tip-KABLA"].get().strip()
        laying = self._form_values["Način polaganja"].get().strip()
        voltage = self._form_values["U"].get()
        drop_key = self._form_values["Ключ ΔU"].get()
        medium_key = self._medium_selected_key

        pi = self._parse_float(self._form_values["Pi, W"].get(), "Pi, W")
        if pi is None:
            return
        kj = self._parse_float(self._form_values["Kj"].get(), "Kj")
        if kj is None:
            return
        eta = self._parse_float(self._form_values["η"].get(), "η")
        if eta is None:
            return
        if eta <= 0 or eta > 1:
            messagebox.showerror(
                "Ошибка ввода",
                "Поле 'η' должно содержать значение в диапазоне (0; 1].",
            )
            logging.error("Недопустимое значение КПД: %s", eta)
            return
        cos_phi = self._parse_float(self._form_values["cos φ"].get(), "cos φ")
        if cos_phi is None:
            return
        if cos_phi <= 0 or abs(cos_phi) > 1:
            messagebox.showerror(
                "Ошибка ввода",
                "Поле 'cos φ' должно содержать значение от 0 (исключительно) до 1.",
            )
            logging.error("Недопустимое значение cosφ: %s", cos_phi)
            return
        length = self._parse_float(self._form_values["Dužina L, m"].get(), "Dužina L, m")
        if length is None:
            return
        area = self._parse_float(self._form_values["Presek, mm²"].get(), "Presek, mm²")
        if area is None or area == 0:
            messagebox.showerror("Ошибка ввода", "Поле 'Presek, mm²' должно быть положительным числом.")
            logging.error("Недопустимое значение сечения: %s", area)
            return

        insulation_meta = self.INSULATION_META.get(insulation_label)
        if insulation_meta is None:
            messagebox.showerror("Ошибка", "Не удалось определить параметры изоляции.")
            return

        loaded_value = self._form_values["Нагруженные жилы (nž)"].get().strip()
        try:
            loaded_cores = int(loaded_value)
            if loaded_cores not in (2, 3):
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка ввода", "Поле 'Нагруженные жилы (nž)' должно быть 2 или 3.")
            logging.error("Некорректное значение нагруженных жил: %s", loaded_value)
            return

        group_value = self._form_values["Кабелей в группе (для S)"].get().strip()
        try:
            circuits_count = int(group_value)
            if circuits_count < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Ошибка ввода",
                "Поле 'Кабелей в группе (для S)' должно содержать целое число от 1 до 20.",
            )
            logging.error("Некорректное значение кабелей в группе: %s", group_value)
            return

        parallel_value = self._form_values["Параллельные кабели (n∥)"].get().strip()
        try:
            n_parallel = int(parallel_value)
            if n_parallel < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Ошибка ввода",
                "Поле 'Параллельные кабели (n∥)' должно содержать целое число от 1 до 6.",
            )
            logging.error("Некорректное значение параллельных кабелей: %s", parallel_value)
            return

        effective_circuits = circuits_count
        if self._consider_parallel_in_s.get():
            effective_circuits = max(1, circuits_count + n_parallel - 1)
        s_coeff = self._lookup_group_factor(effective_circuits)

        temperature = self._try_parse_float(self._form_values["Температура, °C"].get())
        t_coeff = 1.0
        if temperature is not None:
            temp_factor = self._lookup_temperature_factor(insulation_meta["key"], medium_key, temperature)
            if temp_factor is not None and temp_factor > 0:
                t_coeff = temp_factor
            else:
                messagebox.showerror(
                    "Ошибка ввода",
                    "Температура выходит за пределы табличных значений IEC 60364. Укажите корректную температуру.",
                )
                logging.error("Температура вне диапазона: %s °C", temperature)
                return

        pj = pi * kj
        self._form_values["Pj"].set(self._fmt(pj))

        voltage_value = int(voltage)
        phase_factor = 2.0 if loaded_cores == 2 else math.sqrt(3)
        denominator = phase_factor * voltage_value * cos_phi
        if denominator == 0:
            messagebox.showerror("Ошибка расчёта", "Комбинация параметров приводит к делению на ноль.")
            logging.error(
                "Деление на ноль при расчёте тока: pj=%s, eta=%s, voltage=%s, cosφ=%s, phase_factor=%s",
                pj,
                eta,
                voltage_value,
                cos_phi,
                phase_factor,
            )
            return
        icalc_total = (pj / eta) / denominator
        icalc_per_cable = icalc_total / n_parallel

        r_per_km, x_per_km = self._calculate_line_impedance(conductor, insulation_meta["theta"], area, laying)
        r_per_meter = (r_per_km / 1000.0) / n_parallel
        x_per_meter = (x_per_km / 1000.0) / n_parallel

        sin_phi = math.sqrt(max(0.0, 1.0 - min(cos_phi, 1.0) ** 2))
        impedance_drop = r_per_meter * cos_phi + x_per_meter * sin_phi
        delta_u = phase_factor * icalc_total * impedance_drop * length * 100.0 / voltage_value

        limit_delta = self.DROP_LIMIT_KEYS.get(drop_key, 0.0)
        drop_ok = "OK" if delta_u <= limit_delta else "NE"

        base_ampacity = self._lookup_ampacity(insulation_meta["key"], conductor, laying, area, loaded_cores)
        if base_ampacity is None:
            messagebox.showwarning(
                "Предупреждение",
                "Для выбранной комбинации изоляции, проводника и способа прокладки нет табличных данных IEC 60364.\n"
                "Проверка по току пропущена.",
            )
            logging.warning(
                "Нет табличных данных IEC 60364 для комбинации: insulation=%s, conductor=%s, laying=%s, area=%s, cores=%s",
                insulation_meta["key"],
                conductor,
                laying,
                area,
                loaded_cores,
            )
            iz_numeric = None
            compatibility_status = self._("status.no_data")
        else:
            iz_numeric = base_ampacity * s_coeff * t_coeff
            compatibility_status = self._("status.ok")

        if iz_numeric is not None:
            ampacity_ok = "OK" if icalc_per_cable <= iz_numeric else "NE"
        else:
            ampacity_ok = "N/A"

        existing_drop = self._sum_drop_for_circuit(strujni_krug)
        total_drop = existing_drop + delta_u

        in_value = self._try_parse_float(self._form_values["In, A"].get())
        k_value = self._try_parse_float(self._form_values["k"].get())
        i2_value = None
        protection_status = "—"
        if in_value is not None and k_value is not None and in_value > 0 and k_value > 0 and iz_numeric is not None:
            i2_value = in_value * k_value
            iz_total = iz_numeric * n_parallel
            if icalc_total <= in_value <= iz_total and i2_value <= 1.45 * iz_total:
                protection_status = "OK"
            else:
                protection_status = "NE"
        elif in_value is not None and k_value is not None:
            i2_value = in_value * k_value
            protection_status = "N/A"

        if ampacity_ok == "NE" or drop_ok == "NE":
            rec_msg = (
                self._intermediate_vars.get("Рекомендации").get()
                if "Рекомендации" in self._intermediate_vars
                else ""
            )
            if rec_msg and rec_msg.strip() != "—":
                messagebox.showinfo(self._("dialog.recommendations.title"), rec_msg)

        sigma = None
        if conductor in self.RESISTIVITY_20 and self.RESISTIVITY_20[conductor] > 0:
            sigma = 1.0 / self.RESISTIVITY_20[conductor]

        row_data = {
            "Strujni krug": strujni_krug,
            "OD": od,
            "DO": do,
            "E": insulation_label,
            "F": conductor,
            "G": cable,
            "nž": str(loaded_cores),
            "n∥": str(n_parallel),
            "Кабелей в группе (S)": str(circuits_count),
            "Pi": self._fmt(pi),
            "Kj": self._fmt(kj),
            "η": self._fmt(eta, digits=3),
            "Pj": self._fmt(pj),
            "U": voltage,
            "cosφ": self._fmt(cos_phi, digits=3),
            "L": self._fmt(length),
            "Presek": self._fmt(area),
            "Način polaganja": laying,
            "S": self._fmt(s_coeff),
            "T": self._fmt(t_coeff),
            "In [A]": self._fmt(in_value) if in_value is not None else "",
            "k": self._fmt(k_value) if k_value is not None else "",
            "I2 [A]": self._fmt(i2_value) if i2_value is not None else "",
            "Icalc [A]": self._fmt(icalc_total, digits=3),
            "R_base [Ω/km]": self._fmt(r_per_km, digits=3),
            "ϭ": self._fmt(sigma, digits=2) if sigma is not None else "—",
            "Iz [A]": self._fmt(iz_numeric) if iz_numeric is not None else "—",
            "ΔU %": self._fmt(delta_u),
            "Ukupni ΔU %": self._fmt(total_drop),
            "Limit ΔU %": self._fmt(limit_delta),
            "По току": ampacity_ok,
            "По ΔU": drop_ok,
            "Защита": protection_status,
            "Ключ": drop_key,
            "Совместимость IEC": compatibility_status,
        }

        values = [row_data[column] for column in self.TREE_COLUMNS]
        self.tree.insert("", tk.END, values=values)
        self._table_data.append(row_data)

    def clear_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._table_data.clear()
        self._update_intermediate_results()

    def save_project(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Сохранить проект",
            defaultextension=".json",
            filetypes=[("Файл проекта", "*.json"), ("Все файлы", "*.*")],
        )
        if not file_path:
            return

        data = {
            "form": {name: var.get() for name, var in self._form_values.items()},
            "table": self._table_data,
        }

        try:
            with open(file_path, "w", encoding="utf-8") as handle:
                json.dump(data, handle, ensure_ascii=False, indent=2)
        except OSError as exc:
            messagebox.showerror("Ошибка", f"Не удалось сохранить проект: {exc}")
            logging.error("Ошибка сохранения проекта '%s': %s", file_path, exc)
        else:
            messagebox.showinfo("Сохранение", "Проект успешно сохранён.")

    def load_project(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Открыть проект",
            defaultextension=".json",
            filetypes=[("Файл проекта", "*.json"), ("Все файлы", "*.*")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Ошибка", f"Не удалось загрузить проект: {exc}")
            logging.error("Ошибка загрузки проекта '%s': %s", file_path, exc)
            return

        form_data = payload.get("form", {})
        table_data = payload.get("table", [])

        for name, value in form_data.items():
            if name in self._form_values:
                if name == "Среда для Т":
                    self._set_medium_from_value(str(value))
                else:
                    self._form_values[name].set(str(value))
                widget = self._input_widgets.get(name)
                if isinstance(widget, ttk.Combobox):
                    current_values = list(widget.cget("values"))
                    display_value = str(value)
                    if display_value not in current_values and display_value != "":
                        current_values.append(display_value)
                        widget.configure(values=current_values)
                        self._combobox_values[name] = current_values

        self.clear_table()

        for row in table_data:
            if not isinstance(row, dict):
                continue
            normalized = {column: str(row.get(column, "")) for column in self.TREE_COLUMNS}
            self._table_data.append(normalized)
            values = [normalized[column] for column in self.TREE_COLUMNS]
            self.tree.insert("", tk.END, values=values)

        self._update_intermediate_results()

    def export_to_excel(self) -> None:
        if not self._table_data:
            messagebox.showinfo("Экспорт", "Нет данных для экспорта.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not file_path:
            return

        try:
            self._write_workbook(file_path)
            messagebox.showinfo("Экспорт", "Данные успешно сохранены.")
        except (OSError, ValueError) as exc:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {exc}")
            logging.error("Ошибка экспорта Excel '%s': %s", file_path, exc)

    def _write_workbook(self, file_path: str) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Proračuni"

        header_row = [self._(self.TREE_COLUMN_KEYS.get(col, col)) for col in self.TREE_COLUMNS]
        worksheet.append(header_row)

        for row in self._table_data:
            row_values: list[typing.Any] = []
            for column in self.TREE_COLUMNS:
                raw = row.get(column, "")
                if isinstance(raw, (int, float)):
                    row_values.append(raw)
                    continue
                text = str(raw)
                number = self._try_parse_float(text)
                if number is not None and text.strip() not in {"", "—"}:
                    row_values.append(number)
                else:
                    row_values.append(text)
            worksheet.append(row_values)

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for idx, column in enumerate(self.TREE_COLUMNS, start=1):
            max_length = len(header_row[idx - 1])
            for column_cells in worksheet.iter_cols(min_col=idx, max_col=idx, min_row=1, max_row=worksheet.max_row):
                for cell in column_cells:
                    cell_value = cell.value
                    if cell_value is None:
                        continue
                    max_length = max(max_length, len(str(cell_value)))
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 40)

        workbook.save(file_path)

    def _export_to_xlsx_xml(self, file_path: str) -> None:
        workbook_xml = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            "<sheets><sheet name=\"Proračuni\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
            "</workbook>"
        )

        worksheet_xml = self._build_sheet_xml()

        workbook_rels_xml = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            "</Relationships>"
        )

        content_types_xml = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
            "</Types>"
        )

        package_rels_xml = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
            "</Relationships>"
        )

        with ZipFile(file_path, "w") as archive:
            archive.writestr("[Content_Types].xml", content_types_xml)
            archive.writestr("_rels/.rels", package_rels_xml)
            archive.writestr("xl/workbook.xml", workbook_xml)
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)

    def _build_sheet_xml(self) -> str:
        def column_letter(idx: int) -> str:
            result = ""
            while idx:
                idx, remainder = divmod(idx - 1, 26)
                result = chr(65 + remainder) + result
            return result or "A"

        rows: list[str] = []

        header_row = [self._(self.TREE_COLUMN_KEYS.get(col, col)) for col in self.TREE_COLUMNS]
        rows.append(self._build_row_xml(1, header_row, column_letter))

        for row_idx, data in enumerate(self._table_data, start=2):
            row_values = [data[column] for column in self.TREE_COLUMNS]
            rows.append(self._build_row_xml(row_idx, row_values, column_letter))

        sheet_data = "".join(rows)
        worksheet_xml = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            "<sheetData>" + sheet_data + "</sheetData></worksheet>"
        )
        return worksheet_xml

    def _build_row_xml(self, row_number: int, values: list[str], column_letter) -> str:
        cells = []
        for idx, value in enumerate(values, start=1):
            cell_ref = f"{column_letter(idx)}{row_number}"
            value = value or ""
            try:
                float_value = float(value.replace(",", "."))
                if math.isfinite(float_value):
                    cell_xml = f"<c r=\"{cell_ref}\"><v>{float_value}</v></c>"
                else:
                    raise ValueError
            except ValueError:
                cell_xml = (
                    f"<c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{escape(value)}</t></is></c>"
                )
            cells.append(cell_xml)
        return f"<row r=\"{row_number}\">{''.join(cells)}</row>"


def main() -> None:
    app = CableCalcApp()
    app.mainloop()


if __name__ == "__main__":
    main()
