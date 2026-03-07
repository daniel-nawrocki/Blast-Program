import math
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


def _create_header(title_text: str, subtitle_text: str) -> QWidget:
    header = QWidget()
    layout = QVBoxLayout(header)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)

    title = QLabel(title_text)
    title.setObjectName("pageTitle")
    layout.addWidget(title)

    subtitle = QLabel(subtitle_text)
    subtitle.setObjectName("pageSubtitle")
    subtitle.setWordWrap(True)
    layout.addWidget(subtitle)
    return header


def _create_card_button(title_text: str, description: str, action) -> QPushButton:
    button = QPushButton(f"{title_text}\n{description}")
    button.setProperty("variant", "card")
    button.setCursor(Qt.PointingHandCursor)
    button.setMinimumHeight(92)
    button.clicked.connect(action)
    return button


def _create_back_button(label: str, action) -> QPushButton:
    button = QPushButton(label)
    button.setProperty("variant", "secondary")
    button.setCursor(Qt.PointingHandCursor)
    button.clicked.connect(action)
    return button


class StartScreen(QWidget):
    def __init__(self, navigate_to):
        super().__init__()
        self._navigate_to = navigate_to
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(20)

        header = _create_header(
            "Blast Program",
            "Select a module to begin. Each section opens as an independent workspace.",
        )
        layout.addWidget(header)

        cards = QWidget()
        cards_layout = QGridLayout(cards)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        cards_layout.setVerticalSpacing(12)

        for idx, (page_key, title, description) in enumerate(
            [
                ("vibration_tool", "Vibration Tool", "Calculator and site-factor workflows."),
                ("empirical_formula", "Empirical Formula", "Calculate formula from composition inputs."),
                ("diagram_maker", "Diagram Maker", "Build and export process diagrams."),
                ("quick_cheat_sheets", "References/Cheat Sheets", "Open compact references and key values."),
                ("gassing_calculator", "Gassing Calculator", "Workbook-backed calculation workflow."),
            ]
        ):
            button = _create_card_button(
                title,
                description,
                lambda _, target=page_key: self._navigate_to(target),
            )
            row = idx // 2
            col = idx % 2
            cards_layout.addWidget(button, row, col)

        for col in range(2):
            cards_layout.setColumnStretch(col, 1)

        layout.addWidget(cards)
        layout.addStretch()


class VibrationToolScreen(QWidget):
    def __init__(self, navigate_to, navigate_home):
        super().__init__()
        self._navigate_to = navigate_to
        self._navigate_home = navigate_home
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(20)

        header = _create_header(
            "Vibration Tool",
            "Two workflows are available: direct PPV calculation and site-factor calibration.",
        )
        layout.addWidget(header)

        cards = QWidget()
        cards_layout = QGridLayout(cards)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setHorizontalSpacing(12)
        cards_layout.setVerticalSpacing(12)

        cards_layout.addWidget(
            _create_card_button(
                "Vibration Calculator",
                "Estimate PPV using one of 3 standard formulas.",
                lambda: self._navigate_to("vibration_calculator"),
            ),
            0,
            0,
        )
        cards_layout.addWidget(
            _create_card_button(
                "Site Factor Calibrator",
                "Adjust site factor with measured and expected PPV.",
                lambda: self._navigate_to("site_factor_calibrator"),
            ),
            0,
            1,
        )
        cards_layout.setColumnStretch(0, 1)
        cards_layout.setColumnStretch(1, 1)
        layout.addWidget(cards)

        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Start Screen", self._navigate_home))
        footer.addStretch()
        layout.addLayout(footer)


class VibrationCalculatorScreen(QWidget):
    def __init__(self, navigate_back):
        super().__init__()
        self._navigate_back = navigate_back
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(
            "Vibration Calculator",
            "Enter pounds per delay and distance once. The tool reports Dyno and USBM PPV with highest expected value.",
        )
        layout.addWidget(header)

        panel = QWidget()
        panel.setObjectName("contentPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(12)

        form = QFormLayout()
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)

        self._ppd_input = QDoubleSpinBox()
        self._ppd_input.setRange(0.0, 1_000_000_000.0)
        self._ppd_input.setDecimals(2)
        self._ppd_input.setSpecialValueText("")
        self._ppd_input.setValue(0.0)
        self._ppd_input.setSuffix(" lb")

        self._distance_input = QDoubleSpinBox()
        self._distance_input.setRange(0.0, 1_000_000_000.0)
        self._distance_input.setDecimals(2)
        self._distance_input.setSpecialValueText("")
        self._distance_input.setValue(0.0)
        self._distance_input.setSuffix(" ft")

        self._dyno_h_factor_input = QDoubleSpinBox()
        self._dyno_h_factor_input.setRange(1.0, 10000.0)
        self._dyno_h_factor_input.setDecimals(1)
        self._dyno_h_factor_input.setSingleStep(1.0)
        self._dyno_h_factor_input.setValue(160.0)

        form.addRow("Pounds per Delay", self._ppd_input)
        form.addRow("Distance to Seismograph", self._distance_input)
        form.addRow("Dyno Site/Confinement Factor (H)", self._dyno_h_factor_input)
        panel_layout.addLayout(form)

        self._note_label = QLabel(
            "Use H=160 as default, then adjust H using your site confinement or regression-calibrated factor."
        )
        self._note_label.setObjectName("bodyText")
        self._note_label.setWordWrap(True)
        panel_layout.addWidget(self._note_label)

        calculate_button = QPushButton("Calculate PPV")
        calculate_button.clicked.connect(self._calculate)
        panel_layout.addWidget(calculate_button)

        self._result_label = QLabel("PPV result will appear here.")
        self._result_label.setObjectName("bodyText")
        self._result_label.setWordWrap(True)
        panel_layout.addWidget(self._result_label)

        layout.addWidget(panel)
        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Vibration Tool", self._navigate_back))
        footer.addStretch()
        layout.addLayout(footer)

    def _calculate(self) -> None:
        w = self._ppd_input.value()
        d = self._distance_input.value()
        dyno_h_factor = self._dyno_h_factor_input.value()

        if w <= 0.0 or d <= 0.0:
            self._result_label.setText("Enter Pounds per Delay and Distance to Seismograph before calculating.")
            return

        sqrt_scaled_ratio = math.sqrt(w) / d
        sd = d / math.sqrt(w)
        dyno_ppv = dyno_h_factor * (sqrt_scaled_ratio**1.6)
        usbm_best_fit_ppv = 119.0 * (sd**-1.52)
        quarry_upper_bound_ppv = 138.0 * (sd**-1.38)
        average_ppv = (dyno_ppv + usbm_best_fit_ppv) / 2.0
        highest_expected_ppv = quarry_upper_bound_ppv

        self._result_label.setText(
            f"Dyno PPV: {dyno_ppv:.4f} in/sec\n"
            f"USBM PPV: {usbm_best_fit_ppv:.4f} in/sec\n"
            f"Highest Expected PPV: {highest_expected_ppv:.4f} in/sec\n"
            f"Scale Distance: {sd:.4f}"
        )


class SiteFactorCalibratorScreen(QWidget):
    def __init__(self, navigate_back):
        super().__init__()
        self._navigate_back = navigate_back
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(
            "Site Factor Calibrator",
            "Adjust the site factor from measured vibration data using: Adjusted = Current x (Actual / Expected).",
        )
        layout.addWidget(header)

        panel = QWidget()
        panel.setObjectName("contentPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(12)

        form = QFormLayout()
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)

        self._current_factor_input = QDoubleSpinBox()
        self._current_factor_input.setRange(0.01, 10000.0)
        self._current_factor_input.setDecimals(2)
        self._current_factor_input.setValue(160.0)

        self._actual_ppv_input = QDoubleSpinBox()
        self._actual_ppv_input.setRange(0.0, 100.0)
        self._actual_ppv_input.setDecimals(3)
        self._actual_ppv_input.setValue(0.0)
        self._actual_ppv_input.setSuffix(" in/s")

        self._expected_ppv_input = QDoubleSpinBox()
        self._expected_ppv_input.setRange(0.0, 100.0)
        self._expected_ppv_input.setDecimals(3)
        self._expected_ppv_input.setValue(0.0)
        self._expected_ppv_input.setSuffix(" in/s")

        form.addRow("Current Factor", self._current_factor_input)
        form.addRow("Actual PPV", self._actual_ppv_input)
        form.addRow("Expected PPV", self._expected_ppv_input)
        panel_layout.addLayout(form)

        calc_button = QPushButton("Calculate Adjusted Factor")
        calc_button.clicked.connect(self._calculate_adjusted_factor)
        panel_layout.addWidget(calc_button)

        self._result_label = QLabel("Adjusted factor result will appear here.")
        self._result_label.setObjectName("bodyText")
        self._result_label.setWordWrap(True)
        panel_layout.addWidget(self._result_label)

        layout.addWidget(panel)
        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Vibration Tool", self._navigate_back))
        footer.addStretch()
        layout.addLayout(footer)

    def _calculate_adjusted_factor(self) -> None:
        current = self._current_factor_input.value()
        actual = self._actual_ppv_input.value()
        expected = self._expected_ppv_input.value()

        if expected <= 0.0:
            self._result_label.setText("Adjusted factor: enter Expected PPV > 0")
            return

        adjusted = current * (actual / expected)
        self._result_label.setText(f"Adjusted factor: {round(adjusted)}")


class ReferencesScreen(QWidget):
    def __init__(self, navigate_home):
        super().__init__()
        self._navigate_home = navigate_home
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(
            "References/Cheat Sheets",
            "Open field references and quick sheets.",
        )
        layout.addWidget(header)

        panel = QWidget()
        panel.setObjectName("contentPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(10)

        self._status_label = QLabel("Select a document to open.")
        self._status_label.setObjectName("bodyText")
        self._status_label.setWordWrap(True)

        references = [
            (
                "Empirical Cheat Sheet",
                r"c:\Users\danie\OneDrive\Desktop\Work Files\Empirical Cheat Sheet.docx",
            ),
            (
                "Open Pit Book",
                r"c:\Users\danie\OneDrive\Desktop\Work Files\Open Pit Book.pdf",
            ),
        ]

        for label, file_path in references:
            button = QPushButton(label)
            button.clicked.connect(lambda _, p=file_path: self._open_reference(p))
            panel_layout.addWidget(button)

        panel_layout.addWidget(self._status_label)
        layout.addWidget(panel)
        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Start Screen", self._navigate_home))
        footer.addStretch()
        layout.addLayout(footer)

    def _open_reference(self, file_path: str) -> None:
        path = Path(file_path)
        if not path.exists():
            self._status_label.setText(f"File not found: {file_path}")
            return

        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        if not opened:
            self._status_label.setText(f"Could not open: {file_path}")
            return

        self._status_label.setText(f"Opened: {path.name}")


class GassingCalculatorScreen(QWidget):
    WORKBOOK_PATH = Path(r"c:\Users\danie\Downloads\Diff Gassing Titan XL Blends, 02 October 2012, Final.xlsm")
    SHEET_PATH = "xl/worksheets/sheet1.xml"
    SHARED_STRINGS_PATH = "xl/sharedStrings.xml"
    GRID_COLS = [chr(c) for c in range(ord("B"), ord("W") + 1)]
    GRID_ROWS = list(range(3, 67))
    EDITABLE_NUMERIC_CELLS = {
        "P3",
        "C9",
        "C10",
        "C11",
        "C13",
        "C14",
        "D13",
        "D14",
        "E13",
        "E14",
        "F13",
        "F14",
        "C25",
        "U3",
        "W3",
        "W7",
    }

    def __init__(self, navigate_home):
        super().__init__()
        self._navigate_home = navigate_home
        self._input_widgets = {}
        self._template_values = {}
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(
            "Gassing Calculator",
            "Workbook-style native calculator. Layout mirrors the calculation sheet and runs fully in-app.",
        )
        layout.addWidget(header)

        panel = QWidget()
        panel.setObjectName("contentPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(12)

        self._units_combo = QComboBox()
        self._units_combo.addItem("Imperial (T3=0)", 0)
        self._units_combo.addItem("Metric (T3=1)", 1)
        self._units_combo.currentIndexChanged.connect(self._sync_units_cell)
        top_form = QFormLayout()
        top_form.setHorizontalSpacing(16)
        top_form.setVerticalSpacing(8)
        top_form.addRow("T3 Units", self._units_combo)
        self._p3_input = self._make_spin(0.0, 1000.0, 2, 0.0)
        self._c10_input = self._make_spin(-50.0, 200.0, 2, 70.0)
        self._c25_input = self._make_spin(0.0, 10000.0, 3, 1.0)
        self._u3_input = self._make_spin(0.0, 100.0, 4, 0.58)
        self._w3_input = self._make_spin(0.01, 100.0, 4, 14.7)
        self._w7_input = self._make_spin(0.0, 1000.0, 2, 0.0)
        top_form.addRow("P3 Hole Diameter", self._p3_input)
        top_form.addRow("C10 Temperature/Adjustment", self._c10_input)
        top_form.addRow("C25 Target Density", self._c25_input)
        top_form.addRow("U3 Gas Constant", self._u3_input)
        top_form.addRow("W3 Gas Constant", self._w3_input)
        top_form.addRow("W7 Water Deck", self._w7_input)
        panel_layout.addLayout(top_form)

        sheet_title = QLabel("Calculation Sheet (Workbook View)")
        sheet_title.setObjectName("bodyText")
        panel_layout.addWidget(sheet_title)

        self._sheet_table = QTableWidget(len(self.GRID_ROWS), len(self.GRID_COLS))
        self._sheet_table.setAlternatingRowColors(False)
        self._sheet_table.verticalHeader().setVisible(True)
        self._sheet_table.horizontalHeader().setVisible(True)
        self._sheet_table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self._sheet_table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self._sheet_table.verticalHeader().setDefaultSectionSize(28)
        for c in range(len(self.GRID_COLS)):
            self._sheet_table.setColumnWidth(c, 96)
        self._sheet_table.setHorizontalHeaderLabels(self.GRID_COLS)
        self._sheet_table.setVerticalHeaderLabels([str(r) for r in self.GRID_ROWS])
        panel_layout.addWidget(self._sheet_table)

        action_row = QHBoxLayout()
        reload_btn = QPushButton("Reload Workbook Layout")
        reload_btn.setProperty("variant", "secondary")
        reload_btn.clicked.connect(self._load_template_from_workbook)
        action_row.addWidget(reload_btn)
        action_row.addStretch()
        panel_layout.addLayout(action_row)

        calc_btn = QPushButton("Run Gassing Calculation")
        calc_btn.clicked.connect(self._calculate)
        panel_layout.addWidget(calc_btn)

        self._result_label = QLabel("Sheet outputs are shown in the cells above.")
        self._result_label.setObjectName("bodyText")
        self._result_label.setWordWrap(True)
        panel_layout.addWidget(self._result_label)

        self._status_label = QLabel("Native mode active.")
        self._status_label.setObjectName("bodyText")
        self._status_label.setWordWrap(True)
        panel_layout.addWidget(self._status_label)

        layout.addWidget(panel)
        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Start Screen", self._navigate_home))
        footer.addStretch()
        layout.addLayout(footer)

        self._load_template_from_workbook()

    def _make_spin(self, minimum: float, maximum: float, decimals: int, default: float) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(minimum, maximum)
        spin.setDecimals(decimals)
        spin.setValue(default)
        spin.setButtonSymbols(QDoubleSpinBox.NoButtons)
        spin.setAlignment(Qt.AlignRight)
        return spin

    def _cell_index(self, cell_ref: str) -> tuple[int, int]:
        col = "".join(ch for ch in cell_ref if ch.isalpha())
        row = int("".join(ch for ch in cell_ref if ch.isdigit()))
        return self.GRID_ROWS.index(row), self.GRID_COLS.index(col)

    def _set_cell_text(self, cell_ref: str, text: str) -> None:
        if not any(ch.isdigit() for ch in cell_ref):
            return
        row_idx, col_idx = self._cell_index(cell_ref)
        item = self._sheet_table.item(row_idx, col_idx)
        if item is None:
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self._sheet_table.setItem(row_idx, col_idx, item)
        item.setText(text)

    def _get_numeric_from_cell(self, cell_ref: str, default: float = 0.0) -> float:
        widget = self._input_widgets.get(cell_ref)
        if widget is not None:
            return widget.value()

        if cell_ref == "P3":
            return self._p3_input.value()
        if cell_ref == "C10":
            return self._c10_input.value()
        if cell_ref == "C25":
            return self._c25_input.value()
        if cell_ref == "U3":
            return self._u3_input.value()
        if cell_ref == "W3":
            return self._w3_input.value()
        if cell_ref == "W7":
            return self._w7_input.value()

        try:
            row_idx, col_idx = self._cell_index(cell_ref)
        except ValueError:
            return default
        item = self._sheet_table.item(row_idx, col_idx)
        if item is None:
            return default
        try:
            return float(item.text())
        except Exception:
            return default

    def _load_template_from_workbook(self) -> None:
        self._input_widgets.clear()
        for r_idx, row in enumerate(self.GRID_ROWS):
            for c_idx, col in enumerate(self.GRID_COLS):
                cell_ref = f"{col}{row}"
                self._sheet_table.removeCellWidget(r_idx, c_idx)
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if cell_ref in self.EDITABLE_NUMERIC_CELLS:
                    item.setFlags(Qt.ItemIsEnabled)
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self._sheet_table.setItem(r_idx, c_idx, item)

        workbook_values = self._read_workbook_values()
        self._template_values = workbook_values
        if not workbook_values:
            self._status_label.setText(f"Workbook template not loaded: {self.WORKBOOK_PATH}")
            return

        for row in self.GRID_ROWS:
            for col in self.GRID_COLS:
                cell_ref = f"{col}{row}"
                value = workbook_values.get(cell_ref, "")
                row_idx, col_idx = self._cell_index(cell_ref)
                if cell_ref in self.EDITABLE_NUMERIC_CELLS:
                    spin = self._make_spin(-1_000_000.0, 1_000_000.0, 4, 0.0)
                    try:
                        spin.setValue(float(value))
                    except Exception:
                        spin.setValue(0.0)
                    self._sheet_table.setCellWidget(row_idx, col_idx, spin)
                    self._input_widgets[cell_ref] = spin
                else:
                    self._sheet_table.item(row_idx, col_idx).setText(str(value))

        self._sync_units_cell()
        self._status_label.setText("Workbook layout loaded.")

    def _sync_units_cell(self) -> None:
        self._set_cell_text("T3", str(int(self._units_combo.currentData())))

    def _read_workbook_values(self) -> dict:
        if not self.WORKBOOK_PATH.exists():
            return {}
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        try:
            with zipfile.ZipFile(self.WORKBOOK_PATH) as zf:
                shared_strings = []
                if self.SHARED_STRINGS_PATH in zf.namelist():
                    sst = ET.fromstring(zf.read(self.SHARED_STRINGS_PATH))
                    for si in sst.findall(".//x:si", ns):
                        text_parts = [t.text or "" for t in si.findall(".//x:t", ns)]
                        shared_strings.append("".join(text_parts))

                sheet = ET.fromstring(zf.read(self.SHEET_PATH))
                values = {}
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
                            val = "TRUE" if raw == "1" else "FALSE"
                        else:
                            val = raw
                    values[ref] = val
                return values
        except Exception:
            return {}

    def _calculate(self) -> None:
        p3 = self._p3_input.value()
        c25 = self._c25_input.value()
        w3 = self._w3_input.value()
        u3 = self._u3_input.value()
        w7 = self._w7_input.value()
        c10 = self._c10_input.value()
        is_metric = int(self._units_combo.currentData()) == 1

        if p3 <= 0.0 or c25 <= 0.0 or w3 <= 0.0:
            self._status_label.setText("Enter valid values for P3, C25, and W3 before calculating.")
            return

        c9 = self._get_numeric_from_cell("C9")
        c13 = self._get_numeric_from_cell("C13")
        c14 = self._get_numeric_from_cell("C14")
        d13 = self._get_numeric_from_cell("D13")
        d14 = self._get_numeric_from_cell("D14")
        e13 = self._get_numeric_from_cell("E13")
        e14 = self._get_numeric_from_cell("E14")
        f13 = self._get_numeric_from_cell("F13")
        f14 = self._get_numeric_from_cell("F14")

        def safe_div(a: float, b: float) -> float:
            return a / b if b != 0 else 0.0

        o7 = c13 + safe_div(d13 * d14, c14) + safe_div(e13 * e14, c14) + safe_div(f13 * f14, c14) + safe_div(w7, c14)
        q7 = 0.0 if d13 == 0 or d14 == 0 else d13 + safe_div(e13 * e14, d14) + safe_div(f13 * f14, d14) + safe_div(w7, d14)
        s7 = 0.0 if e13 == 0 or e14 == 0 else e13 + safe_div(f13 * f14, e14) + safe_div(w7, e14)
        u7 = 0.0 if f13 == 0 or f14 == 0 else f13 + safe_div(w7, f14)

        def calc_density(x7: float, x13: float, x14: float) -> float:
            if x13 == 0 or x14 == 0:
                return 0.0
            term = (x7 - 0.5 * x13) * u3 / w3
            denom = term + safe_div(c25, x14)
            return ((1.0 + term) / denom) * c25 if denom != 0 else 0.0

        c16 = calc_density(o7, c13, c14)
        d16 = calc_density(q7, d13, d14)
        e16 = calc_density(s7, e13, e14)
        f16 = calc_density(u7, f13, f14)

        mass_coeff = 0.000785 if is_metric else 0.3405
        energy_coeff = 4.186 if is_metric else 0.454

        def calc_mass(x13: float, x14: float, x16: float) -> float:
            if x13 == 0 or x14 == 0:
                return 0.0
            return mass_coeff * (p3**2) * x16 * x13

        c19 = calc_mass(c13, c14, c16)
        d19 = calc_mass(d13, d14, d16)
        e19 = calc_mass(e13, e14, e16)
        f19 = calc_mass(f13, f14, f16)

        def calc_energy(x13: float, x14: float, x19: float) -> float:
            if x13 == 0 or x14 == 0:
                return 0.0
            return energy_coeff * (880.0 - 2.0 * c10) * x19 / x13

        c22 = calc_energy(c13, c14, c19)
        d22 = calc_energy(d13, d14, d19)
        e22 = calc_energy(e13, e14, e19)
        f22 = calc_energy(f13, f14, f19)

        g19 = c19 + d19 + e19 + f19
        total_len = c13 + d13 + e13 + f13
        p28 = ((c16 * c13) + (d16 * d13) + (e16 * e13) + (f16 * f13)) / total_len if total_len != 0 else 0.0
        g20 = p28
        q3 = 1.31 if p3 >= 6 else (-0.0049 * (p3**2) + 0.0587 * p3 + 1.1357)
        r62 = max(0.0, c9 - g19)

        outputs = {
            "O7": o7,
            "Q7": q7,
            "S7": s7,
            "U7": u7,
            "C16": c16,
            "D16": d16,
            "E16": e16,
            "F16": f16,
            "C19": c19,
            "D19": d19,
            "E19": e19,
            "F19": f19,
            "C22": c22,
            "D22": d22,
            "E22": e22,
            "F22": f22,
            "G19": g19,
            "G20": g20,
            "P28": p28,
            "Q3": q3,
            "R62": r62,
        }
        for ref, val in outputs.items():
            self._set_cell_text(ref, f"{val:.4f}")

        self._result_label.setText("Calculation sheet updated.")
        self._status_label.setText("Calculation complete (native mode).")


class PlaceholderScreen(QWidget):
    def __init__(self, title_text: str, description: str, navigate_home):
        super().__init__()
        self._navigate_home = navigate_home
        self._build_ui(title_text, description)

    def _build_ui(self, title_text: str, description: str) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(title_text, "This module is scaffolded and ready for implementation.")
        layout.addWidget(header)

        panel = QWidget()
        panel.setObjectName("contentPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(12)

        body = QLabel(description)
        body.setObjectName("bodyText")
        body.setWordWrap(True)
        panel_layout.addWidget(body)

        layout.addWidget(panel)
        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Start Screen", self._navigate_home))
        footer.addStretch()
        layout.addLayout(footer)
