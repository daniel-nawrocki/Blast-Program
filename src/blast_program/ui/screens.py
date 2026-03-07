import math

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDoubleSpinBox,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
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
                ("quick_cheat_sheets", "Quick Cheat Sheets", "Open compact references and key values."),
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
