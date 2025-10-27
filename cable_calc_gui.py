import json
import math
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape
from zipfile import ZipFile


class CableCalcApp(tk.Tk):
    WINDOW_TITLE = "Proračun kablova"
    WINDOW_GEOMETRY = "1200x800"

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
    TEMPERATURE_MEDIA = ["Воздух", "Грунт"]
    INSTALLATION_METHODS = ["A1", "A2", "B1", "B2", "C", "D", "E", "F", "G"]
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
    )

    def __init__(self) -> None:
        super().__init__()
        self.title(self.WINDOW_TITLE)
        self.geometry(self.WINDOW_GEOMETRY)

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

        self._form_values: dict[str, tk.Variable] = {}
        self._input_widgets: dict[str, ttk.Widget] = {}
        self._input_styles: dict[str, str] = {}
        self._intermediate_vars: dict[str, tk.StringVar] = {}
        self._intermediate_labels: dict[str, ttk.Label] = {}
        self._table_data: list[dict[str, str]] = []
        self._last_temperature_warning: tuple[str, str, float] | None = None

        self._build_menu()
        self._build_layout()

    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Сохранить проект…", command=self.save_project)
        file_menu.add_command(label="Загрузить проект…", command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label="Экспорт в Excel…", command=self.export_to_excel)
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.config(menu=menubar)

    def _build_layout(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        main_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="Расчёт")

        help_tab = ttk.Frame(notebook)
        notebook.add(help_tab, text="Помощь")

        container = ttk.Frame(main_tab)
        container.pack(fill=tk.BOTH, expand=True)

        form_frame = ttk.LabelFrame(container, text="Ввод данных")
        form_frame.pack(fill=tk.X, expand=False, side=tk.TOP, pady=(0, 10))
        self._build_form(form_frame)

        intermediate_frame = ttk.LabelFrame(container, text="Промежуточные результаты")
        intermediate_frame.pack(fill=tk.X, expand=False, side=tk.TOP, pady=(0, 10))
        self._build_intermediate_panel(intermediate_frame)

        table_frame = ttk.LabelFrame(container, text="Результаты расчёта")
        table_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        self._build_table(table_frame)

        self._register_form_traces()

        self._build_help_tab(help_tab)

    def _build_help_tab(self, parent: ttk.Frame) -> None:
        text = tk.Text(parent, wrap="word", height=10)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        help_lines = [
            "Описание коэффициентов IEC 60364:\n",
            "S — коэффициент группировки кабелей. Он учитывает влияние совместной прокладки нескольких кабелей на допустимый ток. Значение < 1 уменьшает допустимый ток при плотной укладке.\n",
            "T — коэффициент температуры окружающей среды. Корректирует допустимую нагрузку в зависимости от фактической температуры воздуха или грунта относительно табличных условий.\n",
            "η — КПД установки. При η < 1 ток уменьшается, если Pi учитывает потери, и увеличивается, если Pi описывает полезную мощность. Укажите значение согласно паспорту оборудования.\n",
            "Kj — коэффициент спроса (коэффициент одновременности) для расчёта нагрузки группы потребителей.\n",
            "ΔU — допустимое падение напряжения по выбранному ключу (UIDM, SVDM, SVTS, UITS) согласно разделам IEC 60364, указывающее максимально допустимое отклонение напряжения в процентах.\n",
            "cos φ — коэффициент мощности нагрузки.\n",
            "При необходимости уточнения коэффициентов IEC 60364 применяйте значения из национальных приложений или таблиц стандарта, учитывая условия прокладки и категорию потребителей.\n",
        ]

        text.insert("1.0", "\n".join(help_lines))
        text.configure(state="disabled")

    def _build_form(self, parent: ttk.Frame) -> None:
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
            ("Число цепей", "1"),
            ("Среда для Т", self.TEMPERATURE_MEDIA[0]),
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

            ttk.Label(grid, text=label).grid(row=row, column=label_col, sticky=tk.W, pady=4, padx=(0, 8))

            var = tk.StringVar(value=default)
            self._form_values[label] = var

            if label == "Tip-IZOLACIJE":
                widget = ttk.Combobox(grid, textvariable=var, values=self.INSULATION_OPTIONS, state="readonly")
            elif label == "Tip-PROVODNIKA":
                widget = ttk.Combobox(grid, textvariable=var, values=self.CONDUCTOR_TYPES, state="readonly")
            elif label == "U":
                widget = ttk.Combobox(grid, textvariable=var, values=self.VOLTAGE_LEVELS, state="readonly")
            elif label == "Način polaganja":
                widget = ttk.Combobox(grid, textvariable=var, values=self.INSTALLATION_METHODS, state="readonly")
            elif label == "Нагруженные жилы (nž)":
                widget = ttk.Combobox(grid, textvariable=var, values=["2", "3"], state="readonly")
            elif label == "Число цепей":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=[str(i) for i in range(1, 21)], state="readonly"
                )
            elif label == "Среда для Т":
                widget = ttk.Combobox(grid, textvariable=var, values=self.TEMPERATURE_MEDIA, state="readonly")
            elif label == "Ключ ΔU":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=list(self.DROP_LIMIT_KEYS.keys()), state="readonly"
                )
            elif label in {"Pj", "S", "T"}:
                widget = ttk.Entry(grid, textvariable=var, state="readonly")
            else:
                widget = ttk.Entry(grid, textvariable=var)

            widget.grid(row=row, column=entry_col, sticky=tk.EW, pady=4)
            widget_class = widget.winfo_class()
            original_style = widget.cget("style") or widget_class
            self._input_widgets[label] = widget
            self._input_styles[label] = original_style

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
            ("I2 [A]", "I2 [A]"),
            ("Защита", "Защита"),
        ]

        columns = 3
        for col in range(columns * 2):
            weight = 1 if col % 2 == 1 else 0
            grid.columnconfigure(col, weight=weight)

        for index, (label_text, key) in enumerate(specs):
            row = index // columns
            label_col = (index % columns) * 2
            value_col = label_col + 1

            ttk.Label(grid, text=label_text, style="ResultKey.TLabel").grid(
                row=row, column=label_col, sticky=tk.W, pady=4, padx=(0, 8)
            )

            var = tk.StringVar(value="—")
            value_label = ttk.Label(grid, textvariable=var, style="ResultValue.TLabel")
            value_label.grid(row=row, column=value_col, sticky=tk.EW, pady=4)
            self._intermediate_vars[key] = var
            self._intermediate_labels[key] = value_label

        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(button_frame, text="Рассчитать и добавить строку", command=self.add_row).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        ttk.Button(button_frame, text="Очистить список", command=self.clear_table).pack(side=tk.LEFT)

    def _build_table(self, parent: ttk.Frame) -> None:
        columns = self.TREE_COLUMNS

        tree_container = ttk.Frame(parent)
        tree_container.pack(fill=tk.BOTH, expand=True)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        tree = ttk.Treeview(tree_container, columns=columns, show="headings")
        self.tree = tree

        for col in columns:
            tree.heading(col, text=col)
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
            return None
        try:
            return float(value.replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка ввода", f"Поле '{field_name}' содержит недопустимое значение: {value}")
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
        if medium == "Грунт":
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

    def _show_temperature_warning(self, insulation_key: str, medium: str, temperature: float) -> None:
        rounded_temp = round(temperature, 1)
        key = (insulation_key, medium, rounded_temp)
        if self._last_temperature_warning == key:
            return
        self._last_temperature_warning = key
        messagebox.showwarning(
            "Температура вне диапазона",
            "Для выбранной изоляции и среды отсутствует табличный коэффициент при температуре "
            f"{rounded_temp} °C. Проверьте корректность условий или используйте значения в пределах таблиц IEC 60364.",
        )

    def _update_intermediate_results(self, *_: object) -> None:
        if not self._intermediate_vars:
            return

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
        circuits_value = self._form_values["Число цепей"].get().strip()
        medium = self._form_values["Среда для Т"].get().strip()
        if medium not in self.TEMPERATURE_MEDIA:
            medium = self.TEMPERATURE_MEDIA[0]
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
            circuits_count = int(circuits_value)
            if circuits_count < 1:
                raise ValueError
        except ValueError:
            circuits_count = 1
            circuits_alert = True
        self._set_entry_alert("Число цепей", circuits_alert)

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
            temp_factor = self._lookup_temperature_factor(insulation_meta["key"], medium, temperature)
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
            if temperature is not None and insulation_meta is not None:
                self._show_temperature_warning(insulation_meta["key"], medium, temperature)
        self._form_values["T"].set(t_display)
        self._intermediate_vars["T"].set(t_display or "—")
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

        icalc = None
        phase_factor = None
        if pj is not None and cos_phi is not None and voltage_value and eta_coeff is not None:
            phase_factor = 2.0 if loaded_cores == 2 else math.sqrt(3)
            denominator = phase_factor * voltage_value * cos_phi
            if denominator:
                icalc = (pj * eta_coeff) / denominator
                self._intermediate_vars["Icalc [A]"].set(f"{icalc:.3f}")
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
        if area is not None and insulation_meta is not None:
            base_ampacity = self._lookup_ampacity(
                insulation_meta["key"], conductor, laying, area, loaded_cores
            )

        iz_numeric = None
        if base_ampacity is not None and t_coeff is not None:
            iz_numeric = base_ampacity * s_coeff * t_coeff
            self._intermediate_vars["Iz [A]"].set(f"{iz_numeric:.2f}")
        elif base_ampacity is None:
            self._intermediate_vars["Iz [A]"].set("—")

        ampacity_status = None
        if base_ampacity is None:
            ampacity_status = "N/A"
        elif iz_numeric is None or icalc is None:
            ampacity_status = "—"
        else:
            ampacity_status = "OK" if icalc <= iz_numeric else "NE"
        ampacity_alert = ampacity_status == "NE" or area_alert
        if ampacity_status is not None:
            self._intermediate_vars["По току"].set(ampacity_status)
            self._set_result_alert("По току", ampacity_alert)
        self._set_entry_alert("Presek, mm²", ampacity_alert)

        delta_u = None
        if (
            icalc is not None
            and phase_factor is not None
            and length is not None
            and cos_phi is not None
            and r_per_km is not None
            and x_per_km is not None
            and voltage_value
        ):
            r_per_meter = r_per_km / 1000.0
            x_per_meter = x_per_km / 1000.0
            sin_phi = math.sqrt(max(0.0, 1.0 - min(1.0, cos_phi) ** 2))
            impedance_drop = r_per_meter * cos_phi + x_per_meter * sin_phi
            delta_u = phase_factor * icalc * impedance_drop * length * 100.0 / voltage_value
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

        if iz_numeric is None or icalc is None or in_value is None or i2_value is None:
            if in_value is None or i2_value is None:
                protection_status = "—"
            else:
                protection_status = "N/A"
        else:
            within_nominal = icalc <= in_value <= iz_numeric
            overload_check = i2_value <= 1.45 * iz_numeric
            if within_nominal and overload_check:
                protection_status = "OK"
            else:
                protection_status = "NE"
                protection_alert = True

        self._intermediate_vars["Защита"].set(protection_status)
        self._set_result_alert("Защита", protection_alert)
        self._set_entry_alert("In, A", protection_alert)

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
        medium = self._form_values["Среда для Т"].get().strip()

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
            return
        cos_phi = self._parse_float(self._form_values["cos φ"].get(), "cos φ")
        if cos_phi is None:
            return
        if cos_phi <= 0 or abs(cos_phi) > 1:
            messagebox.showerror(
                "Ошибка ввода",
                "Поле 'cos φ' должно содержать значение от 0 (исключительно) до 1.",
            )
            return
        length = self._parse_float(self._form_values["Dužina L, m"].get(), "Dužina L, m")
        if length is None:
            return
        area = self._parse_float(self._form_values["Presek, mm²"].get(), "Presek, mm²")
        if area is None or area == 0:
            messagebox.showerror("Ошибка ввода", "Поле 'Presek, mm²' должно быть положительным числом.")
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
            return

        circuits_value = self._form_values["Число цепей"].get().strip()
        try:
            circuits_count = int(circuits_value)
            if circuits_count < 1:
                raise ValueError
        except ValueError:
            circuits_count = 1

        s_coeff = self._lookup_group_factor(circuits_count)

        temperature = self._try_parse_float(self._form_values["Температура, °C"].get())
        t_coeff = 1.0
        if temperature is not None:
            temp_factor = self._lookup_temperature_factor(insulation_meta["key"], medium, temperature)
            if temp_factor is not None and temp_factor > 0:
                t_coeff = temp_factor
            else:
                messagebox.showerror(
                    "Ошибка ввода",
                    "Температура выходит за пределы табличных значений IEC 60364. Укажите корректную температуру.",
                )
                return

        pj = pi * kj
        self._form_values["Pj"].set(f"{pj:.2f}")

        voltage_value = int(voltage)
        phase_factor = 2.0 if loaded_cores == 2 else math.sqrt(3)
        denominator = phase_factor * voltage_value * cos_phi
        if denominator == 0:
            messagebox.showerror("Ошибка расчёта", "Комбинация параметров приводит к делению на ноль.")
            return
        icalc = (pj * eta) / denominator

        r_per_km, x_per_km = self._calculate_line_impedance(conductor, insulation_meta["theta"], area, laying)
        r_per_meter = r_per_km / 1000.0
        x_per_meter = x_per_km / 1000.0

        sin_phi = math.sqrt(max(0.0, 1.0 - min(cos_phi, 1.0) ** 2))
        impedance_drop = r_per_meter * cos_phi + x_per_meter * sin_phi
        delta_u = phase_factor * icalc * impedance_drop * length * 100.0 / voltage_value

        limit_delta = self.DROP_LIMIT_KEYS.get(drop_key, 0.0)
        drop_ok = "OK" if delta_u <= limit_delta else "NE"

        base_ampacity = self._lookup_ampacity(insulation_meta["key"], conductor, laying, area, loaded_cores)
        if base_ampacity is None:
            messagebox.showwarning(
                "Предупреждение",
                "Для выбранной комбинации изоляции, проводника и способа прокладки нет табличных данных IEC 60364.\n"
                "Проверка по току пропущена.",
            )
            iz_numeric = None
        else:
            iz_numeric = base_ampacity * s_coeff * t_coeff

        if iz_numeric is not None:
            ampacity_ok = "OK" if icalc <= iz_numeric else "NE"
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
            if icalc <= in_value <= iz_numeric and i2_value <= 1.45 * iz_numeric:
                protection_status = "OK"
            else:
                protection_status = "NE"
        elif in_value is not None and k_value is not None:
            i2_value = in_value * k_value
            protection_status = "N/A"

        row_data = {
            "Strujni krug": strujni_krug,
            "OD": od,
            "DO": do,
            "E": insulation_label,
            "F": conductor,
            "G": cable,
            "nž": str(loaded_cores),
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
            "Icalc [A]": f"{icalc:.3f}",
            "R_base [Ω/km]": f"{r_per_km:.3f}",
            "Iz [A]": iz_display,
            "ΔU %": f"{delta_u:.2f}",
            "Ukupni ΔU %": f"{total_drop:.2f}",
            "Limit ΔU %": f"{limit_delta:.2f}",
            "По току": ampacity_ok,
            "По ΔU": drop_ok,
            "Защита": protection_status,
            "Ключ": drop_key,
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
            return

        form_data = payload.get("form", {})
        table_data = payload.get("table", [])

        for name, value in form_data.items():
            if name in self._form_values:
                self._form_values[name].set(str(value))

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

        header_row = self.TREE_COLUMNS
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
