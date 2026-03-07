import re
import zipfile
import xml.etree.ElementTree as ET
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from blast_program.ui.screens import _create_back_button, _create_header


class _FormulaParser:
    def __init__(self, text: str, get_cell, get_range):
        self._tokens = self._tokenize(text)
        self._pos = 0
        self._get_cell = get_cell
        self._get_range = get_range

    def _tokenize(self, text: str):
        text = text.lstrip("=")
        pattern = r'"[^"]*"|<=|>=|<>|[=+\-*/^(),:]|\$?[A-Z]{1,3}\$?\d+|[A-Z_][A-Z0-9_\.]*|\d+\.\d+|\d+'
        return re.findall(pattern, text.replace(" ", ""))

    def _peek(self):
        return self._tokens[self._pos] if self._pos < len(self._tokens) else None

    def _take(self):
        token = self._peek()
        self._pos += 1
        return token

    def _accept(self, value):
        if self._peek() == value:
            self._pos += 1
            return True
        return False

    def parse(self):
        if not self._tokens:
            return ""
        return self._comparison()

    def _comparison(self):
        left = self._add()
        while self._peek() in ("=", "<>", "<", ">", "<=", ">="):
            op = self._take()
            right = self._add()
            left = self._cmp(op, left, right)
        return left

    def _add(self):
        left = self._mul()
        while self._peek() in ("+", "-"):
            op = self._take()
            right = self._mul()
            left = self._num(left) + self._num(right) if op == "+" else self._num(left) - self._num(right)
        return left

    def _mul(self):
        left = self._pow()
        while self._peek() in ("*", "/"):
            op = self._take()
            right = self._pow()
            if op == "*":
                left = self._num(left) * self._num(right)
            else:
                d = self._num(right)
                left = self._num(left) / d if d != 0 else 0.0
        return left

    def _pow(self):
        left = self._unary()
        while self._peek() == "^":
            self._take()
            right = self._unary()
            left = self._num(left) ** self._num(right)
        return left

    def _unary(self):
        if self._accept("+"):
            return self._num(self._unary())
        if self._accept("-"):
            return -self._num(self._unary())
        return self._primary()

    def _primary(self):
        token = self._peek()
        if token is None:
            return ""
        if token == "(":
            self._take()
            value = self._comparison()
            self._accept(")")
            return value
        if token.startswith('"') and token.endswith('"'):
            self._take()
            return token[1:-1]
        if re.fullmatch(r"\d+\.\d+|\d+", token):
            self._take()
            return float(token)
        if token.upper() in ("TRUE", "FALSE"):
            self._take()
            return token.upper() == "TRUE"
        if re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", token):
            ref1 = self._take().replace("$", "")
            if self._accept(":"):
                ref2 = self._take().replace("$", "")
                return self._get_range(ref1, ref2)
            return self._get_cell(ref1)
        if re.fullmatch(r"[A-Z_][A-Z0-9_\.]*", token):
            name = self._take().upper()
            if self._accept("("):
                args = []
                if self._peek() != ")":
                    while True:
                        args.append(self._comparison())
                        if not self._accept(","):
                            break
                self._accept(")")
                return self._call(name, args)
            return name
        self._take()
        return ""

    def _call(self, name, args):
        if name == "IF":
            cond = self._bool(args[0]) if len(args) > 0 else False
            yes = args[1] if len(args) > 1 else ""
            no = args[2] if len(args) > 2 else ""
            return yes if cond else no
        if name == "OR":
            return any(self._bool(a) for a in args)
        if name == "INDEX":
            if len(args) < 2:
                return 0.0
            arr = args[0]
            r = int(self._num(args[1]))
            c = int(self._num(args[2])) if len(args) > 2 else 1
            if not isinstance(arr, list) or not arr:
                return 0.0
            rr, cc = r - 1, c - 1
            if rr < 0 or rr >= len(arr):
                return 0.0
            row = arr[rr]
            if not isinstance(row, list) or cc < 0 or cc >= len(row):
                return 0.0
            return row[cc]
        return 0.0

    def _num(self, value):
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if value in ("", None):
            return 0.0
        try:
            return float(str(value))
        except Exception:
            return 0.0

    def _bool(self, value):
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return value != 0
        if isinstance(value, str):
            s = value.strip().upper()
            if s in ("", "FALSE", "0"):
                return False
            if s in ("TRUE", "1"):
                return True
        return bool(value)

    def _cmp(self, op, left, right):
        if isinstance(left, str) or isinstance(right, str):
            a, b = str(left), str(right)
        else:
            a, b = self._num(left), self._num(right)
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        return False


class GassingCalculatorScreen(QWidget):
    PROJECT_ROOT = Path(__file__).resolve().parents[3]
    WORKBOOK_PATH = Path(
        os.environ.get(
            "BLAST_WORKBOOK_PATH",
            str(PROJECT_ROOT / "assets" / "workbooks" / "Diff Gassing Titan XL Blends, 02 October 2012, Final.xlsm"),
        )
    )
    SHEET_PATH = "xl/worksheets/sheet1.xml"
    SHARED_STRINGS_PATH = "xl/sharedStrings.xml"

    def __init__(self, navigate_home):
        super().__init__()
        self._navigate_home = navigate_home
        self._template_values = {}
        self._formula_map = {}
        self._chart_compact = None
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(
            "Gassing Calculator",
            "Spreadsheet-style layout with limited inputs and workbook-matched outputs.",
        )
        layout.addWidget(header)

        panel = QWidget()
        panel.setObjectName("contentPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(12)

        # Inputs
        self._units_combo = QComboBox()
        self._units_combo.addItem("Imperial", 2)
        self._units_combo.addItem("Metric", 1)

        self._hole_diameter_input = self._make_spin(0.0, 1000.0, 2, 4.5)
        self._hole_depth_input = self._make_spin(0.0, 10000.0, 2, 38.0)

        self._wet_hole_check = QCheckBox("Wet Hole")
        self._alt_needed_check = QCheckBox("Alt Diameter")
        self._alt_diameter_input = self._make_spin(0.0, 1000.0, 2, 5.75)
        self._alt_needed_check.toggled.connect(self._on_alt_toggle)

        self._bottom_column_input = self._make_spin(0.0, 10000.0, 0, 29.0)
        self._mid_bot_column_input = self._make_spin(0.0, 10000.0, 0, 0.0)
        self._mid_top_column_input = self._make_spin(0.0, 10000.0, 0, 0.0)
        self._top_column_input = self._make_spin(0.0, 10000.0, 0, 0.0)

        self._bottom_density_input = self._make_spin(0.0, 5.0, 2, 1.1)
        self._mid_bot_density_input = self._make_spin(0.0, 5.0, 2, 0.0)
        self._mid_top_density_input = self._make_spin(0.0, 5.0, 2, 0.0)
        self._top_density_input = self._make_spin(0.0, 5.0, 2, 0.0)

        # Spreadsheet-like grid
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        grid.setColumnStretch(0, 2)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(2, 1)
        grid.setColumnStretch(3, 1)
        grid.setColumnStretch(4, 1)
        grid.setColumnStretch(5, 1)

        title = QLabel("Titan Calculator (Differential Energy)")
        title.setObjectName("pageSubtitle")
        grid.addWidget(title, 0, 0, 1, 6)

        grid.addWidget(QLabel("Hole Diameter"), 1, 0)
        grid.addWidget(self._hole_diameter_input, 1, 1)
        grid.addWidget(QLabel("Units"), 1, 2)
        grid.addWidget(self._units_combo, 1, 3)
        grid.addWidget(self._wet_hole_check, 1, 4, 1, 2)

        grid.addWidget(QLabel("Hole Depth"), 2, 0)
        grid.addWidget(self._hole_depth_input, 2, 1)
        grid.addWidget(self._alt_needed_check, 2, 2, 1, 2)
        grid.addWidget(QLabel("Alt Diameter Value"), 2, 4)
        grid.addWidget(self._alt_diameter_input, 2, 5)

        grid.addWidget(QLabel("Explosive Column"), 3, 0)
        grid.addWidget(self._bottom_column_input, 3, 1)
        grid.addWidget(self._mid_bot_column_input, 3, 2)
        grid.addWidget(self._mid_top_column_input, 3, 3)
        grid.addWidget(self._top_column_input, 3, 4)

        grid.addWidget(QLabel("Cup Density"), 4, 0)
        grid.addWidget(self._bottom_density_input, 4, 1)
        grid.addWidget(self._mid_bot_density_input, 4, 2)
        grid.addWidget(self._mid_top_density_input, 4, 3)
        grid.addWidget(self._top_density_input, 4, 4)

        self._avg_density_labels = [self._output_label() for _ in range(4)]
        self._bottom_density_labels = [self._output_label() for _ in range(4)]
        self._pounds_labels = [self._output_label() for _ in range(4)]

        grid.addWidget(QLabel("Ave Density"), 5, 0)
        for i, lbl in enumerate(self._avg_density_labels, start=1):
            grid.addWidget(lbl, 5, i)

        grid.addWidget(QLabel("Bottom Density"), 6, 0)
        for i, lbl in enumerate(self._bottom_density_labels, start=1):
            grid.addWidget(lbl, 6, i)

        grid.addWidget(QLabel("Pounds to load"), 7, 0)
        for i, lbl in enumerate(self._pounds_labels, start=1):
            grid.addWidget(lbl, 7, i)

        panel_layout.addLayout(grid)

        action_row = QHBoxLayout()
        reload_btn = QPushButton("Reload Workbook")
        reload_btn.setProperty("variant", "secondary")
        reload_btn.clicked.connect(self._load_template_from_workbook)
        action_row.addWidget(reload_btn)
        action_row.addStretch()
        panel_layout.addLayout(action_row)

        calc_btn = QPushButton("Run Gassing Calculation")
        calc_btn.clicked.connect(self._calculate)
        panel_layout.addWidget(calc_btn)

        self._result_label = QLabel("")
        self._result_label.setObjectName("bodyText")
        self._result_label.setWordWrap(True)
        panel_layout.addWidget(self._result_label)

        self._status_label = QLabel("")
        self._status_label.setObjectName("bodyText")
        self._status_label.setWordWrap(True)
        panel_layout.addWidget(self._status_label)

        self._chart_layout = QGridLayout()
        self._chart_layout.setHorizontalSpacing(12)
        self._chart_layout.setVerticalSpacing(8)

        self._as_loaded_title = QLabel("As Loaded")
        self._as_loaded_bar = QWidget()
        self._as_loaded_bar_layout = QHBoxLayout(self._as_loaded_bar)
        self._as_loaded_bar_layout.setContentsMargins(0, 0, 0, 0)
        self._as_loaded_bar_layout.setSpacing(2)

        self._final_title = QLabel("Final")
        self._final_bar = QWidget()
        self._final_bar_layout = QHBoxLayout(self._final_bar)
        self._final_bar_layout.setContentsMargins(0, 0, 0, 0)
        self._final_bar_layout.setSpacing(2)

        self._rise_title = QLabel("Rise")
        self._rise_value_label = QLabel("0.00")
        self._rise_value_label.setProperty("sheetOutput", True)
        self._rise_value_label.setAlignment(Qt.AlignCenter)
        self._reflow_chart_layout(compact=self.width() < 980)

        panel_layout.addLayout(self._chart_layout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(panel)

        layout.addWidget(scroll, 1)

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Start Screen", self._navigate_home))
        footer.addStretch()
        layout.addLayout(footer)
        layout.setStretch(1, 1)

        self._load_template_from_workbook()
        self._on_alt_toggle(self._alt_needed_check.isChecked())

    def _make_spin(self, minimum: float, maximum: float, decimals: int, default: float) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(minimum, maximum)
        spin.setDecimals(decimals)
        spin.setValue(default)
        spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
        spin.setAlignment(Qt.AlignRight)
        return spin

    def _output_label(self) -> QLabel:
        label = QLabel("")
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        return label

    def _reflow_chart_layout(self, compact: bool) -> None:
        if self._chart_compact == compact:
            return
        self._chart_compact = compact

        while self._chart_layout.count():
            child = self._chart_layout.takeAt(0)
            widget = child.widget()
            if widget is not None:
                widget.setParent(None)

        if compact:
            self._chart_layout.addWidget(self._as_loaded_title, 0, 0)
            self._chart_layout.addWidget(self._as_loaded_bar, 1, 0)
            self._chart_layout.addWidget(self._final_title, 2, 0)
            self._chart_layout.addWidget(self._final_bar, 3, 0)
            self._chart_layout.addWidget(self._rise_title, 4, 0)
            self._chart_layout.addWidget(self._rise_value_label, 5, 0)
        else:
            self._chart_layout.addWidget(self._as_loaded_title, 0, 0)
            self._chart_layout.addWidget(self._as_loaded_bar, 1, 0)
            self._chart_layout.addWidget(self._final_title, 0, 1)
            self._chart_layout.addWidget(self._final_bar, 1, 1)
            self._chart_layout.addWidget(self._rise_title, 0, 2)
            self._chart_layout.addWidget(self._rise_value_label, 1, 2)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._reflow_chart_layout(compact=self.width() < 980)

    @staticmethod
    def _col_to_num(col: str) -> int:
        n = 0
        for ch in col:
            n = n * 26 + (ord(ch) - ord("A") + 1)
        return n

    @staticmethod
    def _num_to_col(n: int) -> str:
        out = []
        while n > 0:
            n, rem = divmod(n - 1, 26)
            out.append(chr(rem + ord("A")))
        return "".join(reversed(out))

    @staticmethod
    def _split_ref(ref: str):
        m = re.fullmatch(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", ref)
        if not m:
            return None
        return bool(m.group(1)), m.group(2), bool(m.group(3)), int(m.group(4))

    def _shift_ref(self, ref: str, dr: int, dc: int) -> str:
        parsed = self._split_ref(ref)
        if not parsed:
            return ref
        abs_col, col, abs_row, row = parsed
        cnum = self._col_to_num(col)
        ncol = cnum if abs_col else max(1, cnum + dc)
        nrow = row if abs_row else max(1, row + dr)
        return f"{'$' if abs_col else ''}{self._num_to_col(ncol)}{'$' if abs_row else ''}{nrow}"

    def _shift_formula(self, formula: str, base_ref: str, target_ref: str) -> str:
        b = self._split_ref(base_ref)
        t = self._split_ref(target_ref)
        if not b or not t:
            return formula
        dr = t[3] - b[3]
        dc = self._col_to_num(t[1]) - self._col_to_num(b[1])
        return re.sub(r"\$?[A-Z]{1,3}\$?\d+", lambda m: self._shift_ref(m.group(0), dr, dc), formula)

    def _read_workbook_template(self):
        if not self.WORKBOOK_PATH.exists():
            return {}, {}
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        values = {}
        formulas = {}
        with zipfile.ZipFile(self.WORKBOOK_PATH) as zf:
            shared_strings = []
            if self.SHARED_STRINGS_PATH in zf.namelist():
                sst = ET.fromstring(zf.read(self.SHARED_STRINGS_PATH))
                for si in sst.findall(".//x:si", ns):
                    parts = [t.text or "" for t in si.findall(".//x:t", ns)]
                    shared_strings.append("".join(parts))

            sheet = ET.fromstring(zf.read(self.SHEET_PATH))
            shared_formula_base = {}
            pending_shared = []
            for c in sheet.findall(".//x:c", ns):
                ref = c.attrib.get("r", "")
                if not ref:
                    continue
                t_attr = c.attrib.get("t", "")
                v_node = c.find("x:v", ns)
                inline_t = c.find(".//x:is/x:t", ns)
                val = ""
                if inline_t is not None and inline_t.text:
                    val = inline_t.text
                elif v_node is not None and v_node.text is not None:
                    raw = v_node.text
                    if t_attr == "s":
                        try:
                            val = shared_strings[int(raw)]
                        except Exception:
                            val = raw
                    elif t_attr == "b":
                        val = raw == "1"
                    else:
                        try:
                            val = float(raw)
                        except Exception:
                            val = raw
                values[ref] = val

                f = c.find("x:f", ns)
                if f is None:
                    continue
                f_text = (f.text or "").strip()
                if f.attrib.get("t") == "shared":
                    si = f.attrib.get("si")
                    if f_text:
                        shared_formula_base[si] = (ref, f_text)
                        formulas[ref] = f_text
                    else:
                        pending_shared.append((ref, si))
                elif f_text:
                    formulas[ref] = f_text

            # Resolve shared formulas in a second pass so order in XML does not matter.
            for ref, si in pending_shared:
                if si in shared_formula_base:
                    b_ref, b_formula = shared_formula_base[si]
                    formulas[ref] = self._shift_formula(b_formula, b_ref, ref)
        return values, formulas

    def _load_template_from_workbook(self) -> None:
        values, formulas = self._read_workbook_template()
        self._template_values = values
        self._formula_map = formulas
        if not values:
            self._status_label.setText(f"Workbook template not loaded: {self.WORKBOOK_PATH}")
            return

        for ref, widget in [
            ("P3", self._hole_diameter_input),
            ("C9", self._hole_depth_input),
            ("E8", self._alt_diameter_input),
            ("C13", self._bottom_column_input),
            ("D13", self._mid_bot_column_input),
            ("E13", self._mid_top_column_input),
            ("F13", self._top_column_input),
            ("C14", self._bottom_density_input),
            ("D14", self._mid_bot_density_input),
            ("E14", self._mid_top_density_input),
            ("F14", self._top_density_input),
        ]:
            try:
                widget.setValue(float(values.get(ref, 0.0) or 0.0))
            except Exception:
                widget.setValue(0.0)

        t3 = int(values.get("T3", 0) or 0)
        self._units_combo.setCurrentIndex(1 if t3 == 1 else 0)
        self._wet_hole_check.setChecked(self._as_bool(values.get("O10", False)))
        self._alt_needed_check.setChecked(self._as_bool(values.get("Q10", False)))
        self._on_alt_toggle(self._alt_needed_check.isChecked())
        self._status_label.setText("")

    def _on_alt_toggle(self, checked: bool) -> None:
        self._alt_diameter_input.setEnabled(checked)

    @staticmethod
    def _as_bool(value) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return value != 0
        if isinstance(value, str):
            return value.strip().upper() in ("TRUE", "1", "YES", "Y")
        return bool(value)

    def _build_runtime_values(self):
        values = dict(self._template_values)
        values["T3"] = int(self._units_combo.currentData())
        values["O10"] = self._wet_hole_check.isChecked()
        values["Q10"] = self._alt_needed_check.isChecked()
        values["P3"] = self._hole_diameter_input.value()
        values["C9"] = self._hole_depth_input.value()
        values["E8"] = self._alt_diameter_input.value()
        values["C13"] = self._bottom_column_input.value()
        values["D13"] = self._mid_bot_column_input.value()
        values["E13"] = self._mid_top_column_input.value()
        values["F13"] = self._top_column_input.value()
        values["C14"] = self._bottom_density_input.value()
        values["D14"] = self._mid_bot_density_input.value()
        values["E14"] = self._mid_top_density_input.value()
        values["F14"] = self._top_density_input.value()
        return values

    @staticmethod
    def _as_float(value) -> float:
        try:
            return float(value)
        except Exception:
            return 0.0

    def _clear_layout(self, layout: QHBoxLayout) -> None:
        while layout.count():
            child = layout.takeAt(0)
            w = child.widget()
            if w is not None:
                w.deleteLater()

    def _draw_stacked_bar(self, layout: QHBoxLayout, values: list[float], colors: list[str], labels: list[str]) -> None:
        self._clear_layout(layout)
        positive = [max(0.0, v) for v in values]
        total = sum(positive)
        if total <= 0:
            empty = QLabel("No data")
            empty.setObjectName("bodyText")
            layout.addWidget(empty)
            return
        chart_col = QWidget()
        chart_col_layout = QVBoxLayout(chart_col)
        chart_col_layout.setContentsMargins(0, 0, 0, 0)
        chart_col_layout.setSpacing(1)
        chart_col_layout.addStretch()

        target_height = 260
        heights = []
        for val in positive:
            if val <= 0:
                heights.append(0)
            else:
                heights.append(max(20, int(round((val / total) * target_height))))

        for idx in reversed(range(len(positive))):
            val = positive[idx]
            if val <= 0:
                continue
            seg = QLabel(f"{labels[idx]}\n{val:.2f}")
            seg.setAlignment(Qt.AlignCenter)
            seg.setFixedHeight(heights[idx])
            seg.setStyleSheet(
                f"background-color: {colors[idx]}; border: 1px solid #ffffff; color: #111827; font-size: 11px;"
            )
            seg.setToolTip(f"{labels[idx]}: {val:.4f}")
            chart_col_layout.addWidget(seg)

        layout.addWidget(chart_col)
        layout.addStretch()

    def _calculate(self) -> None:
        values = self._build_runtime_values()
        runtime_formulas = dict(self._formula_map)

        if not self._alt_needed_check.isChecked():
            runtime_formulas.pop("P3", None)
            values["P3"] = self._hole_diameter_input.value()
            values["Q10"] = False
        else:
            values["Q10"] = True

        cache = {}
        visiting = set()

        def get_cell(ref):
            if ref in cache:
                return cache[ref]
            if ref in visiting:
                return values.get(ref, 0.0)
            if ref in runtime_formulas:
                visiting.add(ref)
                parser = _FormulaParser(runtime_formulas[ref], get_cell, get_range)
                try:
                    result = parser.parse()
                except Exception:
                    result = values.get(ref, 0.0)
                visiting.discard(ref)
                cache[ref] = result
                values[ref] = result
                return result
            return values.get(ref, 0.0)

        def get_range(start_ref, end_ref):
            s = self._split_ref(start_ref)
            e = self._split_ref(end_ref)
            if not s or not e:
                return [[0.0]]
            sc, ec = self._col_to_num(s[1]), self._col_to_num(e[1])
            sr, er = s[3], e[3]
            row_step = 1 if er >= sr else -1
            col_step = 1 if ec >= sc else -1
            out = []
            for r in range(sr, er + row_step, row_step):
                row_vals = []
                for c in range(sc, ec + col_step, col_step):
                    row_vals.append(get_cell(f"{self._num_to_col(c)}{r}"))
                out.append(row_vals)
            return out

        for ref in runtime_formulas.keys():
            get_cell(ref)

        for idx, col in enumerate(["C", "D", "E", "F"]):
            self._avg_density_labels[idx].setText(f"{self._as_float(values.get(f'{col}16', 0.0)):.2f}")
            self._bottom_density_labels[idx].setText(f"{self._as_float(values.get(f'{col}17', 0.0)):.2f}")
            self._pounds_labels[idx].setText(f"{round(self._as_float(values.get(f'{col}19', 0.0)))}")

        as_loaded_values = [
            self._as_float(values.get("U20", 0.0)),
            self._as_float(values.get("U19", 0.0)),
            self._as_float(values.get("U18", 0.0)),
            self._as_float(values.get("U17", 0.0)),
            self._as_float(values.get("U16", 0.0)),
        ]
        if sum(max(0.0, v) for v in as_loaded_values) <= 0:
            # Pre-rise fallback from intermediate column heights.
            as_loaded_values = [
                self._as_float(values.get("R20", 0.0)),
                self._as_float(values.get("R19", 0.0)),
                self._as_float(values.get("R18", 0.0)),
                self._as_float(values.get("R17", 0.0)),
                self._as_float(values.get("R16", 0.0)),
            ]
        if sum(max(0.0, v) for v in as_loaded_values) <= 0:
            # Secondary fallback by unit mode if R-chain is unavailable.
            if int(values.get("T3", 2)) == 1:
                as_loaded_values = [
                    self._as_float(values.get("Q20", 0.0)),
                    self._as_float(values.get("Q19", 0.0)),
                    self._as_float(values.get("Q18", 0.0)),
                    self._as_float(values.get("Q17", 0.0)),
                    self._as_float(values.get("Q16", 0.0)),
                ]
            else:
                as_loaded_values = [
                    self._as_float(values.get("P20", 0.0)),
                    self._as_float(values.get("P19", 0.0)),
                    self._as_float(values.get("P18", 0.0)),
                    self._as_float(values.get("P17", 0.0)),
                    self._as_float(values.get("P16", 0.0)),
                ]
        if sum(max(0.0, v) for v in as_loaded_values) <= 0:
            # Last resort: workbook's currently stored U-cells.
            as_loaded_values = [
                self._as_float(self._template_values.get("U20", 0.0)),
                self._as_float(self._template_values.get("U19", 0.0)),
                self._as_float(self._template_values.get("U18", 0.0)),
                self._as_float(self._template_values.get("U17", 0.0)),
                self._as_float(self._template_values.get("U16", 0.0)),
            ]
        final_values = [
            self._as_float(values.get("C13", 0.0)),
            self._as_float(values.get("D13", 0.0)),
            self._as_float(values.get("E13", 0.0)),
            self._as_float(values.get("F13", 0.0)),
            self._as_float(values.get("S16", 0.0)),
        ]

        colors = ["#ef4444", "#f97316", "#facc15", "#22c55e", "#38bdf8"]
        as_loaded_labels = ["Top", "Mid-Top", "Mid-Bot", "Bottom", "Unloaded Collar"]
        final_labels = ["Bottom", "Mid-Bot", "Mid-Top", "Top", "Unloaded Collar"]
        self._draw_stacked_bar(self._as_loaded_bar_layout, as_loaded_values, colors, as_loaded_labels)
        self._draw_stacked_bar(self._final_bar_layout, final_values, colors, final_labels)
        rise = as_loaded_values[4] - final_values[4]
        self._rise_value_label.setText(f"{rise:.2f}")
        self._as_loaded_bar.update()
        self._final_bar.update()

        hole_avg_density = self._as_float(values.get("G20", 0.0))
        if hole_avg_density == 0.0:
            lengths = [self._as_float(values.get(f"{col}13", 0.0)) for col in ["C", "D", "E", "F"]]
            densities = [self._as_float(values.get(f"{col}16", 0.0)) for col in ["C", "D", "E", "F"]]
            total_length = sum(lengths)
            if total_length > 0:
                hole_avg_density = sum(l * d for l, d in zip(lengths, densities)) / total_length

        total_pounds = self._as_float(values.get("G19", 0.0))
        if total_pounds == 0.0:
            total_pounds = sum(self._as_float(values.get(f"{col}19", 0.0)) for col in ["C", "D", "E", "F"])

        self._result_label.setText(
            f"Hole Average Density: {hole_avg_density:.2f}\n"
            f"Total Pounds for Hole: {total_pounds:.2f}"
        )
        self._status_label.setText("")
