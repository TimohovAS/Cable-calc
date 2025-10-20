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
    INSTALLATION_METHODS = ["A1", "A2", "B1", "B2", "C", "D", "E", "F"]
    DROP_LIMIT_KEYS = {
        "UIDM": 5.0,
        "SVDM": 3.0,
        "SVTS": 5.0,
        "UITS": 8.0,
    }

    RESISTIVITY_20 = {"Cu": 0.017241, "Al": 0.028264}
    TEMP_COEFF = {"Cu": 0.00393, "Al": 0.00403}
    REACTANCE_PER_KM = {"default": 0.08, "D": 0.09}

    AMPACITY_TABLE = {
        "PVC": {
            "Cu": {
                "B1": {
                    1.5: 14.5,
                    2.5: 19.5,
                    4.0: 26.0,
                    6.0: 32.0,
                    10.0: 44.0,
                    16.0: 57.0,
                    25.0: 76.0,
                    35.0: 96.0,
                    50.0: 115.0,
                    70.0: 146.0,
                    95.0: 176.0,
                    120.0: 202.0,
                    150.0: 231.0,
                    185.0: 262.0,
                    240.0: 308.0,
                },
                "B2": {
                    1.5: 16.5,
                    2.5: 22.0,
                    4.0: 29.0,
                    6.0: 37.0,
                    10.0: 51.0,
                    16.0: 68.0,
                    25.0: 89.0,
                    35.0: 110.0,
                    50.0: 131.0,
                    70.0: 164.0,
                    95.0: 196.0,
                    120.0: 225.0,
                    150.0: 257.0,
                    185.0: 293.0,
                    240.0: 344.0,
                },
                "C": {
                    1.5: 18.0,
                    2.5: 24.0,
                    4.0: 32.0,
                    6.0: 41.0,
                    10.0: 57.0,
                    16.0: 76.0,
                    25.0: 101.0,
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
                    10.0: 46.0,
                    16.0: 61.0,
                    25.0: 80.0,
                    35.0: 99.0,
                    50.0: 118.0,
                    70.0: 146.0,
                    95.0: 176.0,
                    120.0: 202.0,
                    150.0: 229.0,
                    185.0: 260.0,
                    240.0: 302.0,
                },
            },
            "Al": {
                "B1": {
                    10.0: 34.0,
                    16.0: 44.0,
                    25.0: 58.0,
                    35.0: 72.0,
                    50.0: 86.0,
                    70.0: 109.0,
                    95.0: 131.0,
                    120.0: 150.0,
                    150.0: 171.0,
                    185.0: 194.0,
                    240.0: 225.0,
                },
                "B2": {
                    10.0: 39.0,
                    16.0: 50.0,
                    25.0: 66.0,
                    35.0: 81.0,
                    50.0: 95.0,
                    70.0: 121.0,
                    95.0: 146.0,
                    120.0: 167.0,
                    150.0: 190.0,
                    185.0: 215.0,
                    240.0: 250.0,
                },
                "C": {
                    10.0: 47.0,
                    16.0: 61.0,
                    25.0: 80.0,
                    35.0: 99.0,
                    50.0: 118.0,
                    70.0: 146.0,
                    95.0: 176.0,
                    120.0: 202.0,
                    150.0: 229.0,
                    185.0: 260.0,
                    240.0: 302.0,
                },
                "D": {
                    16.0: 56.0,
                    25.0: 72.0,
                    35.0: 86.0,
                    50.0: 101.0,
                    70.0: 122.0,
                    95.0: 144.0,
                    120.0: 163.0,
                    150.0: 184.0,
                    185.0: 207.0,
                    240.0: 240.0,
                },
            },
        },
        "XLPE": {
            "Cu": {
                "B1": {
                    1.5: 17.5,
                    2.5: 23.5,
                    4.0: 31.0,
                    6.0: 39.0,
                    10.0: 53.0,
                    16.0: 70.0,
                    25.0: 94.0,
                    35.0: 118.0,
                    50.0: 141.0,
                    70.0: 181.0,
                    95.0: 220.0,
                    120.0: 252.0,
                    150.0: 288.0,
                    185.0: 327.0,
                    240.0: 384.0,
                },
                "B2": {
                    1.5: 19.5,
                    2.5: 26.0,
                    4.0: 34.0,
                    6.0: 44.0,
                    10.0: 61.0,
                    16.0: 81.0,
                    25.0: 109.0,
                    35.0: 137.0,
                    50.0: 165.0,
                    70.0: 206.0,
                    95.0: 249.0,
                    120.0: 286.0,
                    150.0: 327.0,
                    185.0: 372.0,
                    240.0: 437.0,
                },
                "C": {
                    1.5: 23.0,
                    2.5: 30.0,
                    4.0: 41.0,
                    6.0: 53.0,
                    10.0: 73.0,
                    16.0: 98.0,
                    25.0: 129.0,
                    35.0: 159.0,
                    50.0: 195.0,
                    70.0: 250.0,
                    95.0: 306.0,
                    120.0: 353.0,
                    150.0: 407.0,
                    185.0: 467.0,
                    240.0: 545.0,
                },
                "D": {
                    10.0: 55.0,
                    16.0: 73.0,
                    25.0: 96.0,
                    35.0: 118.0,
                    50.0: 140.0,
                    70.0: 174.0,
                    95.0: 210.0,
                    120.0: 241.0,
                    150.0: 273.0,
                    185.0: 309.0,
                    240.0: 360.0,
                },
            },
            "Al": {
                "B1": {
                    10.0: 40.0,
                    16.0: 52.0,
                    25.0: 69.0,
                    35.0: 86.0,
                    50.0: 102.0,
                    70.0: 129.0,
                    95.0: 155.0,
                    120.0: 177.0,
                    150.0: 201.0,
                    185.0: 228.0,
                    240.0: 264.0,
                },
                "B2": {
                    10.0: 45.0,
                    16.0: 58.0,
                    25.0: 77.0,
                    35.0: 95.0,
                    50.0: 112.0,
                    70.0: 142.0,
                    95.0: 173.0,
                    120.0: 198.0,
                    150.0: 225.0,
                    185.0: 255.0,
                    240.0: 296.0,
                },
                "C": {
                    10.0: 55.0,
                    16.0: 73.0,
                    25.0: 96.0,
                    35.0: 118.0,
                    50.0: 140.0,
                    70.0: 174.0,
                    95.0: 210.0,
                    120.0: 241.0,
                    150.0: 273.0,
                    185.0: 309.0,
                    240.0: 360.0,
                },
                "D": {
                    16.0: 67.0,
                    25.0: 87.0,
                    35.0: 104.0,
                    50.0: 122.0,
                    70.0: 149.0,
                    95.0: 177.0,
                    120.0: 200.0,
                    150.0: 225.0,
                    185.0: 253.0,
                    240.0: 292.0,
                },
            },
        },
    }

    TREE_COLUMNS = (
        "Strujni krug",
        "OD",
        "DO",
        "E",
        "F",
        "G",
        "Pi",
        "Kj",
        "Pj",
        "U",
        "cosφ",
        "L",
        "Presek",
        "Način polaganja",
        "S",
        "T",
        "Ucf",
        "Icalc [A]",
        "R_base [Ω/km]",
        "Iz [A]",
        "ΔU %",
        "Limit ΔU %",
        "По току",
        "По ΔU",
        "Ключ",
    )

    def __init__(self) -> None:
        super().__init__()
        self.title(self.WINDOW_TITLE)
        self.geometry(self.WINDOW_GEOMETRY)

        self._form_values: dict[str, tk.Variable] = {}
        self._table_data: list[dict[str, str]] = []

        self._build_menu()
        self._build_layout()

    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Экспорт в Excel…", command=self.export_to_excel)
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.config(menu=menubar)

    def _build_layout(self) -> None:
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        form_frame = ttk.LabelFrame(container, text="Ввод данных")
        form_frame.pack(fill=tk.X, expand=False, side=tk.TOP, pady=(0, 10))
        self._build_form(form_frame)

        table_frame = ttk.LabelFrame(container, text="Результаты расчёта")
        table_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        self._build_table(table_frame)

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
            ("Pj", ""),
            ("U", self.VOLTAGE_LEVELS[0]),
            ("cos φ", ""),
            ("Dužina L, m", ""),
            ("Presek, mm²", ""),
            ("Način polaganja", self.INSTALLATION_METHODS[4]),
            ("S", "1.0"),
            ("T", "1.0"),
            ("Ucf", "1.0"),
            ("Ключ ΔU", list(self.DROP_LIMIT_KEYS.keys())[0]),
        ]

        grid = ttk.Frame(parent)
        grid.pack(fill=tk.X, expand=False, padx=10, pady=10)

        columns = 3
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
            elif label == "Ключ ΔU":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=list(self.DROP_LIMIT_KEYS.keys()), state="readonly"
                )
            elif label == "Pj":
                widget = ttk.Entry(grid, textvariable=var, state="readonly")
            else:
                widget = ttk.Entry(grid, textvariable=var)

            widget.grid(row=row, column=entry_col, sticky=tk.EW, pady=4)

        pi_var = self._form_values["Pi, W"]
        kj_var = self._form_values["Kj"]
        pi_var.trace_add("write", self._update_pj_display)
        kj_var.trace_add("write", self._update_pj_display)
        self._update_pj_display()

        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(button_frame, text="Рассчитать и добавить строку", command=self.add_row).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        ttk.Button(button_frame, text="Очистить список", command=self.clear_table).pack(side=tk.LEFT)

    def _build_table(self, parent: ttk.Frame) -> None:
        columns = self.TREE_COLUMNS
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        self.tree = tree

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor=tk.CENTER)

        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

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

    def _lookup_ampacity(self, insulation_key: str, conductor: str, laying: str, area: float) -> float | None:
        conductor_tables = self.AMPACITY_TABLE.get(insulation_key, {}).get(conductor, {})
        method_table = conductor_tables.get(laying)
        if not method_table:
            return None

        standard_sections = sorted(method_table)
        if not standard_sections:
            return None
        for section in standard_sections:
            if math.isclose(area, section, rel_tol=1e-6, abs_tol=1e-3):
                return method_table[section]

        for lower, upper in zip(standard_sections, standard_sections[1:]):
            if lower <= area <= upper:
                lower_val = method_table[lower]
                upper_val = method_table[upper]
                if upper == lower:
                    return lower_val
                ratio = (area - lower) / (upper - lower)
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
            return
        self._form_values["Pj"].set(f"{pi * kj:.2f}")

    def add_row(self) -> None:
        strujni_krug = self._form_values["Strujni krug"].get().strip()
        od = self._form_values["Deonica OD"].get().strip()
        do = self._form_values["Deonica DO"].get().strip()
        insulation_label = self._form_values["Tip-IZOLACIJE"].get()
        conductor = self._form_values["Tip-PROVODNIKA"].get()
        cable = self._form_values["Oznaka-tip-KABLA"].get().strip()

        pi = self._parse_float(self._form_values["Pi, W"].get(), "Pi, W")
        if pi is None:
            return
        kj = self._parse_float(self._form_values["Kj"].get(), "Kj")
        if kj is None:
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
        s_coeff = self._parse_float(self._form_values["S"].get(), "S")
        if s_coeff is None:
            return
        t_coeff = self._parse_float(self._form_values["T"].get(), "T")
        if t_coeff is None:
            return
        u_coeff = self._parse_float(self._form_values["Ucf"].get(), "Ucf")
        if u_coeff is None:
            return

        laying = self._form_values["Način polaganja"].get().strip()
        voltage = self._form_values["U"].get()
        drop_key = self._form_values["Ключ ΔU"].get()

        insulation_meta = self.INSULATION_META.get(insulation_label)
        if insulation_meta is None:
            messagebox.showerror("Ошибка", "Не удалось определить параметры изоляции.")
            return

        pj = pi * kj
        self._form_values["Pj"].set(f"{pj:.2f}")

        voltage_value = int(voltage)
        if voltage_value == 230:
            icalc = pj / (voltage_value * cos_phi)
            phase_factor = 2.0
        else:
            icalc = pj / (math.sqrt(3) * voltage_value * cos_phi)
            phase_factor = math.sqrt(3)

        r_per_km, x_per_km = self._calculate_line_impedance(conductor, insulation_meta["theta"], area, laying)
        r_per_meter = r_per_km / 1000.0
        x_per_meter = x_per_km / 1000.0

        sin_phi = math.sqrt(max(0.0, 1.0 - min(cos_phi, 1.0) ** 2))
        impedance_drop = r_per_meter * cos_phi + x_per_meter * sin_phi
        delta_u = phase_factor * icalc * impedance_drop * length * 100.0 / voltage_value

        limit_delta = self.DROP_LIMIT_KEYS.get(drop_key, 0.0)
        drop_ok = "OK" if delta_u <= limit_delta else "NE"

        base_ampacity = self._lookup_ampacity(insulation_meta["key"], conductor, laying, area)
        if base_ampacity is None:
            messagebox.showwarning(
                "Предупреждение",
                "Для выбранной комбинации изоляции, проводника и способа прокладки нет табличных данных IEC 60364.\n"
                "Проверка по току пропущена.",
            )
            iz_numeric = None
        else:
            iz_numeric = base_ampacity * s_coeff * t_coeff * u_coeff

        if iz_numeric is not None:
            ampacity_ok = "OK" if icalc <= iz_numeric else "NE"
            iz_display = f"{iz_numeric:.2f}"
        else:
            ampacity_ok = "N/A"
            iz_display = ""

        row_data = {
            "Strujni krug": strujni_krug,
            "OD": od,
            "DO": do,
            "E": insulation_label,
            "F": conductor,
            "G": cable,
            "Pi": f"{pi:.2f}",
            "Kj": f"{kj:.2f}",
            "Pj": f"{pj:.2f}",
            "U": voltage,
            "cosφ": f"{cos_phi:.3f}",
            "L": f"{length:.2f}",
            "Presek": f"{area:.2f}",
            "Način polaganja": laying,
            "S": f"{s_coeff:.2f}",
            "T": f"{t_coeff:.2f}",
            "Ucf": f"{u_coeff:.2f}",
            "Icalc [A]": f"{icalc:.3f}",
            "R_base [Ω/km]": f"{r_per_km:.3f}",
            "Iz [A]": iz_display,
            "ΔU %": f"{delta_u:.2f}",
            "Limit ΔU %": f"{limit_delta:.2f}",
            "По току": ampacity_ok,
            "По ΔU": drop_ok,
            "Ключ": drop_key,
        }

        values = [row_data[column] for column in self.TREE_COLUMNS]
        self.tree.insert("", tk.END, values=values)
        self._table_data.append(row_data)

    def clear_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._table_data.clear()

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
