import math
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape
from zipfile import ZipFile


class CableCalcApp(tk.Tk):
    WINDOW_TITLE = "Proračun kablova"
    WINDOW_GEOMETRY = "1200x800"

    INSULATION_TYPES = ["PVC", "XLP-EPR", "EPR", "PE"]
    CONDUCTOR_TYPES = ["Cu", "Al"]
    VOLTAGE_LEVELS = ["230", "400"]
    DROP_LIMIT_KEYS = {
        "UIDM": 5.0,
        "SVDM": 3.0,
        "SVTS": 5.0,
        "UITS": 8.0,
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
        "R_base [A]",
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
        container = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        form_frame = ttk.Frame(container)
        self._build_form(form_frame)
        container.add(form_frame, weight=1)

        table_frame = ttk.Frame(container)
        self._build_table(table_frame)
        container.add(table_frame, weight=3)

    def _build_form(self, parent: ttk.Frame) -> None:
        field_specs = [
            ("Strujni krug", ""),
            ("Deonica OD", ""),
            ("Deonica DO", ""),
            ("Tip-IZOLACIJE", self.INSULATION_TYPES[0]),
            ("Tip-PROVODNIKA", self.CONDUCTOR_TYPES[0]),
            ("Oznaka-tip-KABLA", ""),
            ("Pi, W", ""),
            ("Kj", ""),
            ("Pj", ""),
            ("U", self.VOLTAGE_LEVELS[0]),
            ("cos φ", ""),
            ("Dužina L, m", ""),
            ("Presek, mm²", ""),
            ("Način polaganja", ""),
            ("S", "1.0"),
            ("T", "1.0"),
            ("Ucf", "1.0"),
            ("Ключ ΔU", list(self.DROP_LIMIT_KEYS.keys())[0]),
        ]

        grid = ttk.Frame(parent)
        grid.pack(fill=tk.BOTH, expand=True)

        for row, (label, default) in enumerate(field_specs):
            ttk.Label(grid, text=label).grid(row=row, column=0, sticky=tk.W, pady=3, padx=(0, 8))

            var = tk.StringVar(value=default)
            self._form_values[label] = var

            if label == "Tip-IZOLACIJE":
                widget = ttk.Combobox(grid, textvariable=var, values=self.INSULATION_TYPES, state="readonly")
            elif label == "Tip-PROVODNIKA":
                widget = ttk.Combobox(grid, textvariable=var, values=self.CONDUCTOR_TYPES, state="readonly")
            elif label == "U":
                widget = ttk.Combobox(grid, textvariable=var, values=self.VOLTAGE_LEVELS, state="readonly")
            elif label == "Ключ ΔU":
                widget = ttk.Combobox(
                    grid, textvariable=var, values=list(self.DROP_LIMIT_KEYS.keys()), state="readonly"
                )
            elif label == "Pj":
                widget = ttk.Entry(grid, textvariable=var, state="readonly")
            else:
                widget = ttk.Entry(grid, textvariable=var)

            widget.grid(row=row, column=1, sticky=tk.EW, pady=3)
            grid.columnconfigure(1, weight=1)

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
        insulation = self._form_values["Tip-IZOLACIJE"].get()
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

        pj = pi * kj
        self._form_values["Pj"].set(f"{pj:.2f}")

        voltage_value = int(voltage)
        if voltage_value == 230:
            icalc = pi / (voltage_value * cos_phi)
        else:
            icalc = pj / (math.sqrt(3) * voltage_value * cos_phi)

        gamma = 56 if conductor == "Cu" else 34
        if voltage_value == 400:
            delta_u = (100.0 * length * pj) / (gamma * area * voltage_value**2)
        else:
            delta_u = (200.0 * length * pj) / (gamma * area * voltage_value**2)

        limit_delta = self.DROP_LIMIT_KEYS.get(drop_key, 0.0)
        drop_ok = "OK" if delta_u <= limit_delta else "NE"

        iz_value = ""
        ampacity_ok = "N/A"

        row_data = {
            "Strujni krug": strujni_krug,
            "OD": od,
            "DO": do,
            "E": insulation,
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
            "R_base [A]": "",
            "Iz [A]": iz_value,
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
