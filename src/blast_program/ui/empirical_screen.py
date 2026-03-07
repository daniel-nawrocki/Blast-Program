import math

from PySide6.QtWidgets import (
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from blast_program.ui.screens import _create_back_button, _create_header


class EmpiricalFormulaScreen(QWidget):
    def __init__(self, navigate_home):
        super().__init__()
        self._navigate_home = navigate_home
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 30, 36, 30)
        layout.setSpacing(18)

        header = _create_header(
            "Empirical Formula Calculation",
            "Pattern footage empirical calculator (ported from empirical-calculator).",
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

        self._rock_type = QComboBox()
        self._rock_type.addItems(["Granite/Hard Limestone", "Soft Limestone/Shale/Sandstone"])

        self._face_height = QDoubleSpinBox()
        self._face_height.setRange(0.0, 10000.0)
        self._face_height.setDecimals(2)
        self._face_height.setValue(30.0)
        self._face_height.setSuffix(" ft")

        self._dh = QDoubleSpinBox()
        self._dh.setRange(0.0, 1000.0)
        self._dh.setDecimals(2)
        self._dh.setValue(6.5)
        self._dh.setSuffix(" in")

        self._pattern_type = QComboBox()
        self._pattern_type.addItems(["Square", "Rectangular"])

        self._initiation = QComboBox()
        self._initiation.addItems(["Top and Bottom", "Bottom"])

        form.addRow("Rock Type", self._rock_type)
        form.addRow("Face Height", self._face_height)
        form.addRow("Hole Diameter Dh", self._dh)
        form.addRow("Pattern Type", self._pattern_type)
        form.addRow("Initiation", self._initiation)
        panel_layout.addLayout(form)

        calc_button = QPushButton("Calculate Empirical Outputs")
        calc_button.clicked.connect(self._calculate)
        panel_layout.addWidget(calc_button)

        self._warning_label = QLabel("")
        self._warning_label.setObjectName("bodyText")
        self._warning_label.setWordWrap(True)
        panel_layout.addWidget(self._warning_label)

        self._results_label = QLabel("Run calculation to view burden, spacing, and stemming.")
        self._results_label.setObjectName("bodyText")
        self._results_label.setWordWrap(True)
        panel_layout.addWidget(self._results_label)

        layout.addWidget(panel)
        layout.addStretch()

        footer = QHBoxLayout()
        footer.addWidget(_create_back_button("Back to Start Screen", self._navigate_home))
        footer.addStretch()
        layout.addLayout(footer)

    @staticmethod
    def _get_empirical_band_and_constant(ratio_r: float, rock_type: str):
        granite = {"A": 1200, "B": 906, "C": 806, "D": 484, "E": 282}
        limestone = {"A": 1560, "B": 1177, "C": 1047, "D": 629, "E": 366}
        constants = granite if rock_type == "Granite/Hard Limestone" else limestone
        if ratio_r >= 13.23:
            return "A", constants["A"], False
        if ratio_r >= 9.45:
            return "B", constants["B"], False
        if ratio_r >= 4.8:
            return "C", constants["C"], False
        if ratio_r >= 2.62:
            return "D", constants["D"], False
        if ratio_r >= 1.84:
            return "E", constants["E"], False
        return "E", constants["E"], True

    @staticmethod
    def _rectangular_k(rock_type: str) -> float:
        return 0.85 if rock_type == "Granite/Hard Limestone" else 0.93

    @staticmethod
    def _stemming_k(initiation: str, band: str) -> float:
        if band == "E":
            return 0.70
        return 0.50 if initiation == "Bottom" else 0.70

    def _calculate(self) -> None:
        rock_type = self._rock_type.currentText()
        face_height = self._face_height.value()
        dh = self._dh.value()
        pattern_type = self._pattern_type.currentText()
        initiation = self._initiation.currentText()

        if dh <= 0.0 or face_height <= 0.0:
            self._warning_label.setText("Face height and Dh must be greater than 0.")
            self._results_label.setText("Enter valid inputs, then run calculation.")
            return

        ratio_r = face_height / dh
        band, constant, ratio_below_range = self._get_empirical_band_and_constant(ratio_r, rock_type)
        pattern_footage = (dh / 12.0) ** 2 * constant
        base = math.sqrt(pattern_footage)
        burden = base if pattern_type == "Square" else self._rectangular_k(rock_type) * base
        spacing = pattern_footage / burden if burden > 0 else 0.0
        subdrill = 0.30 * burden
        stemming = self._stemming_k(initiation, band) * burden

        self._warning_label.setText(
            "Ratio below empirical table range; Band E used as fallback."
            if ratio_below_range
            else ""
        )
        self._results_label.setText(
            f"Burden B: {burden:.3f} ft\n"
            f"Spacing S: {spacing:.3f} ft\n"
            f"Subdrill J (0.30B): {subdrill:.3f} ft\n"
            f"Stemming: {stemming:.3f} ft\n"
            f"R = Height / Dh: {ratio_r:.3f}\n"
            f"Band: {band}\n"
            f"Empirical Constant: {constant}\n"
            f"Pattern Footage (PF): {pattern_footage:.3f} ft^2"
        )
