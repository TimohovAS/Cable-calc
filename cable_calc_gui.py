import json
import logging
import math
import typing
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape
from zipfile import ZipFile


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
        "app.title": {"ru": "Proračun kablova", "sr": "Proračun kablova", "en": "Cable calculation"},
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
        "label.circuit": {"ru": "Strujni krug", "sr": "Strujni krug", "en": "Circuit"},
        "label.segment_from": {"ru": "Deonica OD", "sr": "Deonica OD", "en": "Segment from"},
        "label.segment_to": {"ru": "Deonica DO", "sr": "Deonica DO", "en": "Segment to"},
        "label.insulation": {"ru": "Tip-IZOLACIJE", "sr": "Tip izolacije", "en": "Insulation type"},
        "label.conductor": {"ru": "Tip-PROVODNIKA", "sr": "Tip provodnika", "en": "Conductor type"},
        "label.cable": {"ru": "Oznaka-tip-KABLA", "sr": "Oznaka/tip kabla", "en": "Cable designation"},
        "label.pi": {"ru": "Pi, W", "sr": "Pi, W", "en": "Pi, W"},
        "label.kj": {"ru": "Kj", "sr": "Kj", "en": "Kj"},
        "label.eta": {"ru": "η", "sr": "η", "en": "η"},
        "label.pj": {"ru": "Pj", "sr": "Pj", "en": "Pj"},
        "label.voltage": {"ru": "U", "sr": "U", "en": "U"},
        "label.cos_phi": {"ru": "cos φ", "sr": "cos φ", "en": "cos φ"},
        "label.length": {"ru": "Dužina L, m", "sr": "Dužina L, m", "en": "Length L, m"},
        "label.area": {"ru": "Presek, mm²", "sr": "Presek, mm²", "en": "Cross-section, mm²"},
        "label.installation": {"ru": "Način polaganja", "sr": "Način polaganja", "en": "Installation method"},
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
        "label.medium": {"ru": "Среда для Т", "sr": "Okruženje za T", "en": "Medium for T"},
        "label.temperature": {"ru": "Температура, °C", "sr": "Temperatura, °C", "en": "Temperature, °C"},
        "label.s": {"ru": "S", "sr": "S", "en": "S"},
        "label.t": {"ru": "T", "sr": "T", "en": "T"},
        "label.in": {"ru": "In, A", "sr": "In, A", "en": "In, A"},
        "label.k": {"ru": "k", "sr": "k", "en": "k"},
        "label.drop_key": {"ru": "Ключ ΔU", "sr": "Ključ ΔU", "en": "ΔU key"},
        "label.result.pj": {"ru": "Pj, W", "sr": "Pj, W", "en": "Pj, W"},
        "label.result.icalc": {"ru": "Icalc [A]", "sr": "Icalc [A]", "en": "Icalc [A]"},
        "label.result.rbase": {"ru": "R_base [Ω/km]", "sr": "R_base [Ω/km]", "en": "R_base [Ω/km]"},
        "label.result.iz": {"ru": "Iz [A]", "sr": "Iz [A]", "en": "Iz [A]"},
        "label.result.s": {"ru": "S", "sr": "S", "en": "S"},
        "label.result.t": {"ru": "T", "sr": "T", "en": "T"},
        "label.result.delta": {"ru": "ΔU %", "sr": "ΔU %", "en": "ΔU %"},
        "label.result.total_delta": {"ru": "Ukupni ΔU %", "sr": "Ukupni ΔU %", "en": "Total ΔU %"},
        "label.result.limit_delta": {"ru": "Limit ΔU %", "sr": "Limit ΔU %", "en": "Limit ΔU %"},
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
            "ru": "Совместимость IEC",
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
        "column.circuit": {"ru": "Strujni krug", "sr": "Strujni krug", "en": "Circuit"},
        "column.from": {"ru": "OD", "sr": "OD", "en": "From"},
        "column.to": {"ru": "DO", "sr": "DO", "en": "To"},
        "column.insulation": {"ru": "E", "sr": "E", "en": "E"},
        "column.conductor": {"ru": "F", "sr": "F", "en": "F"},
        "column.cable": {"ru": "G", "sr": "G", "en": "G"},
        "column.cores": {"ru": "nž", "sr": "nž", "en": "nž"},
        "column.n_parallel": {"ru": "n∥", "sr": "n∥", "en": "n∥"},
        "column.group_for_s": {
            "ru": "Кабелей в группе (S)",
            "sr": "Kablovi u grupi (S)",
            "en": "Group cables (S)",
        },
        "column.pi": {"ru": "Pi", "sr": "Pi", "en": "Pi"},
        "column.kj": {"ru": "Kj", "sr": "Kj", "en": "Kj"},
        "column.eta": {"ru": "η", "sr": "η", "en": "η"},
        "column.pj": {"ru": "Pj", "sr": "Pj", "en": "Pj"},
        "column.voltage": {"ru": "U", "sr": "U", "en": "U"},
        "column.cos": {"ru": "cosφ", "sr": "cosφ", "en": "cosφ"},
        "column.length": {"ru": "L", "sr": "L", "en": "L"},
        "column.area": {"ru": "Presek", "sr": "Presek", "en": "Area"},
        "column.installation": {
            "ru": "Način polaganja",
            "sr": "Način polaganja",
            "en": "Installation",
        },
        "column.s": {"ru": "S", "sr": "S", "en": "S"},
        "column.t": {"ru": "T", "sr": "T", "en": "T"},
        "column.in": {"ru": "In [A]", "sr": "In [A]", "en": "In [A]"},
        "column.k": {"ru": "k", "sr": "k", "en": "k"},
        "column.i2": {"ru": "I2 [A]", "sr": "I2 [A]", "en": "I2 [A]"},
        "column.icalc": {"ru": "Icalc [A]", "sr": "Icalc [A]", "en": "Icalc [A]"},
        "column.rbase": {"ru": "R_base [Ω/km]", "sr": "R_base [Ω/km]", "en": "R_base [Ω/km]"},
        "column.iz": {"ru": "Iz [A]", "sr": "Iz [A]", "en": "Iz [A]"},
        "column.drop": {"ru": "ΔU %", "sr": "ΔU %", "en": "ΔU %"},
        "column.total_drop": {
            "ru": "Ukupni ΔU %",
            "sr": "Ukupni ΔU %",
            "en": "Total ΔU %",
        },
        "column.limit_drop": {
            "ru": "Limit ΔU %",
            "sr": "Limit ΔU %",
            "en": "Limit ΔU %",
        },
        "column.ampacity": {"ru": "По току", "sr": "Po struji", "en": "By current"},
        "column.drop_status": {"ru": "По ΔU", "sr": "Po ΔU", "en": "By ΔU"},
        "column.protection": {"ru": "Защита", "sr": "Zaštita", "en": "Protection"},
        "column.key": {"ru": "Ключ", "sr": "Ključ", "en": "Key"},
        "column.compatibility": {
            "ru": "Совместимость IEC",
            "sr": "IEC kompatibilnost",
            "en": "IEC compatibility",
        },
        "column.medium": {"ru": "Среда", "sr": "Okruženje", "en": "Medium"},
        "column.limit": {"ru": "Limit ΔU %", "sr": "Limit ΔU %", "en": "Limit ΔU %"},
        "status.ok": {"ru": "OK", "sr": "OK", "en": "OK"},
        "status.fail": {"ru": "NE", "sr": "NE", "en": "NO"},
        "status.na": {"ru": "N/A", "sr": "N/A", "en": "N/A"},
        "status.no_data": {"ru": "Нет данных", "sr": "Nema podataka", "en": "No data"},
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
        "ru": """Описание коэффициентов IEC 60364:\n\n"
        "S — коэффициент группировки кабелей (Kn). Учитывает взаимное нагревание при совместной прокладке.\n"
        "T — температурный коэффициент (Kt) для воздуха или грунта. Используйте табличные значения IEC 60364-5-52.\n"
        "η — КПД установки. Определяется по паспорту оборудования.\n"
        "Kj — коэффициент спроса (одновременности) для группы потребителей.\n"
        "ΔU — допустимое падение напряжения по выбранному ключу (UIDM, SVDM, SVTS, UITS).\n"
        "cos φ — коэффициент мощности нагрузки.\n"
        "In, k — параметры защитного устройства: номинальный ток и коэффициент I2/In.\n"
        "Проверяйте, что Ib ≤ In ≤ Iz и I2 ≤ 1.45×Iz в соответствии с IEC 60364-4-43.\n""",
        "sr": """Opis koeficijenata prema IEC 60364:\n\n"
        "S — koeficijent grupisanja kablova (Kn) koji uzima u obzir međusobno zagrevanje.\n"
        "T — temperaturni koeficijent (Kt) za vazduh ili tlo prema tabelama IEC 60364-5-52.\n"
        "η — efikasnost postrojenja prema podacima proizvođača.\n"
        "Kj — koeficijent istovremenosti potrošača.\n"
        "ΔU — dozvoljeni pad napona po odabranom ključu (UIDM, SVDM, SVTS, UITS).\n"
        "cos φ — faktor snage opterećenja.\n"
        "In, k — parametri zaštitnog uređaja: nazivna struja i odnos I2/In.\n"
        "Proverite da Ib ≤ In ≤ Iz i da je I2 ≤ 1.45×Iz u skladu sa IEC 60364-4-43.\n""",
        "en": """Description of IEC 60364 factors:\n\n"
        "S – cable grouping factor (Kn) accounting for mutual heating.\n"
        "T – ambient temperature factor (Kt) for air or soil from IEC 60364-5-52 tables.\n"
        "η – installation efficiency as specified by the manufacturer.\n"
        "Kj – demand (diversity) factor for the load group.\n"
        "ΔU – permitted voltage drop according to the selected key (UIDM, SVDM, SVTS, UITS).\n"
        "cos φ – load power factor.\n"
        "In, k – protective device parameters: rated current and I2/In ratio.\n"
        "Ensure Ib ≤ In ≤ Iz and I2 ≤ 1.45×Iz in line with IEC 60364-4-43.\n""",
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
    VOLTAGE_LEVELS = ["230", "400"]
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
    REACTANCE_PER_KM = {"default": 0.08, "D": 0.09}

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
        self._medium_selected_key = (
            self.DEFAULT_MEDIUM
            if self.DEFAULT_MEDIUM in self.TEMPERATURE_MEDIA
            else next(iter(self.TEMPERATURE_MEDIA))
        )
        self._medium_combobox: ttk.Combobox | None = None
        self._help_text_widget: tk.Text | None = None
        self._notebook: ttk.Notebook | None = None
        self._tabs: dict[str, ttk.Frame] = {}
        self._last_icalc: float | None = None

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
            ("Pj", ""),
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
            ("S", "1.0"),
            ("T", "1.0"),
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
            elif label in {"Pj", "S", "T"}:
                widget = ttk.Entry(grid, textvariable=var, state="readonly")
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

        pi_var = self._form_values["Pi, W"]
        kj_var = self._form_values["Kj"]
        pi_var.trace_add("write", self._update_pj_display)
        kj_var.trace_add("write", self._update_pj_display)

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

        add_button = ttk.Button(button_frame, text="", command=self.add_row)
        add_button.pack(side=tk.LEFT, padx=(0, 5))
        self._bind_text(lambda value, widget=add_button: widget.configure(text=value), "button.add_row")

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
            return 0.0, self.REACTANCE_PER_KM.get(laying, self.REACTANCE_PER_KM["default"])

        rho_theta = rho_20 * (1.0 + alpha * (insulation_temp - 20.0))
        if area <= 0:
            r_per_km = 0.0
        else:
            r_per_km = (rho_theta / area) * 1000.0

        x_per_km = self.REACTANCE_PER_KM.get(laying, self.REACTANCE_PER_KM["default"])
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
        current_per_cable = icalc_total / parallel_count

        for a in self.STANDARD_SECTIONS:
            if a < (current_area or 0):
                continue
            iz_base = self._lookup_ampacity(insulation_key, conductor, method, a, loaded_cores)
            if not iz_base:
                continue
            iz_one = iz_base * S * T
            if iz_one < current_per_cable:
                continue
            d = self._drop_pct(
                U,
                cos_phi,
                L,
                conductor,
                insulation_theta,
                a,
                method,
                loaded_cores,
                parallel_count,
                icalc_total,
            )
            if limit_pct is None or d <= limit_pct:
                recs.append(
                    f"Увеличить сечение до {a} мм² → Iz_tot≈{iz_one * parallel_count:.0f} A, ΔU≈{d:.2f}%"
                )
                return recs

        for m in self.METHOD_PREFERENCE:
            if m == method:
                continue
            for a in self.STANDARD_SECTIONS:
                iz_base = self._lookup_ampacity(insulation_key, conductor, m, a, loaded_cores)
                if not iz_base:
                    continue
                iz_one = iz_base * S * T
                if iz_one < current_per_cable:
                    continue
                d = self._drop_pct(
                    U,
                    cos_phi,
                    L,
                    conductor,
                    insulation_theta,
                    a,
                    m,
                    loaded_cores,
                    parallel_count,
                    icalc_total,
                )
                if limit_pct is None or d <= limit_pct:
                    recs.append(
                        f"Сменить метод на {m} и сечение {a} мм² → Iz_tot≈{iz_one * parallel_count:.0f} A, ΔU≈{d:.2f}%"
                    )
                    return recs

        if insulation_key == "PVC":
            xlpe_meta = self.INSULATION_META.get("XLPE/EPR (90°C)")
            xlpe_theta = xlpe_meta["theta"] if xlpe_meta else 90.0
            for a in self.STANDARD_SECTIONS:
                iz_base = self._lookup_ampacity("XLPE", conductor, method, a, loaded_cores)
                if not iz_base:
                    continue
                iz_one = iz_base * S * T
                if iz_one < current_per_cable:
                    continue
                d = self._drop_pct(
                    U,
                    cos_phi,
                    L,
                    conductor,
                    xlpe_theta,
                    a,
                    method,
                    loaded_cores,
                    parallel_count,
                    icalc_total,
                )
                if limit_pct is None or d <= limit_pct:
                    recs.append(
                        f"Перейти на XLPE и {a} мм² → Iz_tot≈{iz_one * parallel_count:.0f} A, ΔU≈{d:.2f}%"
                    )
                    return recs

        if limit_pct is not None:
            a0 = max(current_area or self.STANDARD_SECTIONS[0], self.STANDARD_SECTIONS[0])
            d_now = self._drop_pct(
                U,
                cos_phi,
                L,
                conductor,
                insulation_theta,
                a0,
                method,
                loaded_cores,
                parallel_count,
                icalc_total,
            )
            if d_now > max(limit_pct, 1e-9):
                denom = max(limit_pct, 1e-9)
                n_needed = max(parallel_count, math.ceil(d_now * parallel_count / denom))
                if n_needed > parallel_count:
                    iz_base = self._lookup_ampacity(insulation_key, conductor, method, a0, loaded_cores)
                    if iz_base:
                        iz_one = iz_base * S * T
                        iz_total_display = f"{iz_one * n_needed:.0f}"
                    else:
                        iz_total_display = self._("status.no_data")
                    reduced_drop = d_now * parallel_count / n_needed
                    recs.append(
                        f"Разделить на {n_needed} параллельных кабеля {a0} мм² (метод {method}) → Iz_tot≈{iz_total_display} A, ΔU≈{reduced_drop:.2f}%"
                    )

        recs.append("Снизить число кабелей в группе (для увеличения S) или повысить напряжение.")
        return recs

    def _update_pj_display(self, *_: object) -> None:
        pi = self._try_parse_float(self._form_values["Pi, W"].get())
        kj = self._try_parse_float(self._form_values["Kj"].get())
        if pi is None or kj is None:
            self._form_values["Pj"].set("")
            self._update_intermediate_results()
            return
        self._form_values["Pj"].set(f"{pi * kj:.2f}")
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
            self._intermediate_vars["Limit ΔU %"].set(f"{limit_delta:.2f}")
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

        group_factor = self._lookup_group_factor(circuits_count)
        self._form_values["S"].set(f"{group_factor:.2f}")
        self._intermediate_vars["S"].set(f"{group_factor:.2f}")
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
            t_display = f"{t_coeff:.2f}"
            self._last_temperature_warning = None
        else:
            t_display = ""
            if (
                temperature is not None
                and insulation_meta is not None
                and not self._temperature_editing
            ):
                self._show_temperature_warning(insulation_meta["key"], medium_key, temperature)
        self._form_values["T"].set(t_display)
        self._intermediate_vars["T"].set(t_display or "—")
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

        pj = None
        if pi is not None and kj is not None:
            pj = pi * kj
            self._intermediate_vars["Pj, W"].set(f"{pj:.2f}")

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
                self._intermediate_vars["Icalc [A]"].set(f"{icalc_total:.3f}")
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
            self._intermediate_vars["R_base [Ω/km]"].set(f"{r_per_km:.3f}")

        base_ampacity = None
        compat_display = "—"
        compat_alert = False
        if area is not None and insulation_meta is not None:
            base_ampacity = self._lookup_ampacity(
                insulation_meta["key"], conductor, laying, area, loaded_cores
            )
            if base_ampacity is None:
                compat_display = self._("status.no_data")
                compat_alert = True
            else:
                compat_display = self._("status.ok")
        elif insulation_meta is not None:
            compat_display = "—"

        if "Совместимость IEC" in self._intermediate_vars:
            self._intermediate_vars["Совместимость IEC"].set(compat_display)
            self._set_result_alert("Совместимость IEC", compat_alert)

        iz_one = None
        if base_ampacity is not None and t_coeff is not None:
            iz_one = base_ampacity * s_coeff * t_coeff
            self._intermediate_vars["Iz [A]"].set(f"{iz_one:.2f}")
        elif base_ampacity is None:
            self._intermediate_vars["Iz [A]"].set("—")

        in_range_value = "—"
        if icalc_total is not None and iz_one is not None:
            iz_total = iz_one * n_parallel
            in_range_value = f"{icalc_total:.2f} – {iz_total:.2f}"
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
            self._intermediate_vars["ΔU %"].set(f"{delta_u:.2f}")

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
            self._intermediate_vars["Ukupni ΔU %"].set(f"{total_drop:.2f}")
            if limit_delta is not None:
                total_status = "OK" if total_drop <= limit_delta else "NE"
                self._set_result_alert("Ukupni ΔU %", total_status == "NE")
            else:
                self._set_result_alert("Ukupni ΔU %", False)
        elif existing_drop > 0:
            self._intermediate_vars["Ukupni ΔU %"].set(f"{existing_drop:.2f}")

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
            self._intermediate_vars["I2 [A]"].set(f"{i2_value:.2f}")
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

        s_coeff = self._lookup_group_factor(circuits_count)

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
        self._form_values["Pj"].set(f"{pj:.2f}")

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
            iz_display = f"{iz_numeric:.2f}"
        else:
            ampacity_ok = "N/A"
            iz_display = ""

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
            "Pi": f"{pi:.2f}",
            "Kj": f"{kj:.2f}",
            "η": f"{eta:.3f}",
            "Pj": f"{pj:.2f}",
            "U": voltage,
            "cosφ": f"{cos_phi:.3f}",
            "L": f"{length:.2f}",
            "Presek": f"{area:.2f}",
            "Način polaganja": laying,
            "S": f"{s_coeff:.2f}",
            "T": f"{t_coeff:.2f}",
            "In [A]": f"{in_value:.2f}" if in_value is not None else "",
            "k": f"{k_value:.2f}" if k_value is not None else "",
            "I2 [A]": f"{i2_value:.2f}" if i2_value is not None else "",
            "Icalc [A]": f"{icalc_total:.3f}",
            "R_base [Ω/km]": f"{r_per_km:.3f}",
            "Iz [A]": iz_display,
            "ΔU %": f"{delta_u:.2f}",
            "Ukupni ΔU %": f"{total_drop:.2f}",
            "Limit ΔU %": f"{limit_delta:.2f}",
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
            self._write_simple_xlsx(file_path)
            messagebox.showinfo("Экспорт", "Данные успешно сохранены.")
        except OSError as exc:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {exc}")
            logging.error("Ошибка экспорта Excel '%s': %s", file_path, exc)

    def _write_simple_xlsx(self, file_path: str) -> None:
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
