import math

from PySide6.QtWidgets import (
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QGridLayout,
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
        panel_layout = QHBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(16)

        inputs_col = QWidget()
        inputs_layout = QVBoxLayout(inputs_col)
        inputs_layout.setContentsMargins(0, 0, 0, 0)
        inputs_layout.setSpacing(10)

        inputs_title = QLabel("Inputs")
        inputs_title.setObjectName("pageSubtitle")
        inputs_layout.addWidget(inputs_title)

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

        form.addRow("Rock Type", self._rock_type)
        form.addRow("Face Height", self._face_height)
        form.addRow("Hole Diameter Dh", self._dh)
        form.addRow("Pattern Type", self._pattern_type)
        inputs_layout.addLayout(form)

        calc_button = QPushButton("Calculate Empirical Outputs")
        calc_button.clicked.connect(self._calculate)
        inputs_layout.addWidget(calc_button)

        self._warning_label = QLabel("")
        self._warning_label.setObjectName("bodyText")
        self._warning_label.setWordWrap(True)
        inputs_layout.addWidget(self._warning_label)
        inputs_layout.addStretch()

        results_col = QWidget()
        results_layout = QVBoxLayout(results_col)
        results_layout.setContentsMargins(0, 0, 0, 0)
        results_layout.setSpacing(10)

        results_title = QLabel("Results")
        results_title.setObjectName("pageSubtitle")
        results_layout.addWidget(results_title)

        hero_grid = QGridLayout()
        hero_grid.setHorizontalSpacing(18)
        hero_grid.setVerticalSpacing(4)
        self._burden_value = QLabel("--")
        self._burden_value.setObjectName("pageTitle")
        self._spacing_value = QLabel("--")
        self._spacing_value.setObjectName("pageTitle")
        hero_grid.addWidget(QLabel("Burden B (ft)"), 0, 0)
        hero_grid.addWidget(QLabel("Spacing S (ft)"), 0, 1)
        hero_grid.addWidget(self._burden_value, 1, 0)
        hero_grid.addWidget(self._spacing_value, 1, 1)
        results_layout.addLayout(hero_grid)

        details_grid = QGridLayout()
        details_grid.setHorizontalSpacing(14)
        details_grid.setVerticalSpacing(6)
        self._ratio_value = QLabel("--")
        self._band_value = QLabel("--")
        self._constant_value = QLabel("--")
        self._pf_value = QLabel("--")
        self._subdrill_value = QLabel("--")
        self._stemming_value = QLabel("--")
        details_grid.addWidget(QLabel("R = Height / Dh"), 0, 0)
        details_grid.addWidget(self._ratio_value, 0, 1)
        details_grid.addWidget(QLabel("Band"), 1, 0)
        details_grid.addWidget(self._band_value, 1, 1)
        details_grid.addWidget(QLabel("Empirical Constant"), 2, 0)
        details_grid.addWidget(self._constant_value, 2, 1)
        details_grid.addWidget(QLabel("Pattern Footage (PF)"), 3, 0)
        details_grid.addWidget(self._pf_value, 3, 1)
        details_grid.addWidget(QLabel("Subdrill J (0.30B)"), 4, 0)
        details_grid.addWidget(self._subdrill_value, 4, 1)
        details_grid.addWidget(QLabel("Stemming (Top and Bottom)"), 5, 0)
        details_grid.addWidget(self._stemming_value, 5, 1)
        results_layout.addLayout(details_grid)
        results_layout.addStretch()

        panel_layout.addWidget(inputs_col, 1)
        panel_layout.addWidget(results_col, 1)

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
    def _stemming_k(band: str) -> float:
        if band == "E":
            return 0.70
        return 0.70

    def _calculate(self) -> None:
        rock_type = self._rock_type.currentText()
        face_height = self._face_height.value()
        dh = self._dh.value()
        pattern_type = self._pattern_type.currentText()
        if dh <= 0.0 or face_height <= 0.0:
            self._warning_label.setText("Face height and Dh must be greater than 0.")
            self._burden_value.setText("--")
            self._spacing_value.setText("--")
            self._ratio_value.setText("--")
            self._band_value.setText("--")
            self._constant_value.setText("--")
            self._pf_value.setText("--")
            self._subdrill_value.setText("--")
            self._stemming_value.setText("--")
            return

        ratio_r = face_height / dh
        band, constant, ratio_below_range = self._get_empirical_band_and_constant(ratio_r, rock_type)
        pattern_footage = (dh / 12.0) ** 2 * constant
        base = math.sqrt(pattern_footage)
        burden = base if pattern_type == "Square" else self._rectangular_k(rock_type) * base
        spacing = pattern_footage / burden if burden > 0 else 0.0
        subdrill = 0.30 * burden
        stemming = self._stemming_k(band) * burden

        self._warning_label.setText(
            "Ratio below empirical table range; Band E used as fallback."
            if ratio_below_range
            else ""
        )
        self._burden_value.setText(f"{burden:.3f}")
        self._spacing_value.setText(f"{spacing:.3f}")
        self._ratio_value.setText(f"{ratio_r:.3f}")
        self._band_value.setText(band)
        self._constant_value.setText(str(constant))
        self._pf_value.setText(f"{pattern_footage:.3f} ft^2")
        self._subdrill_value.setText(f"{subdrill:.3f} ft")
        self._stemming_value.setText(f"{stemming:.3f} ft")
