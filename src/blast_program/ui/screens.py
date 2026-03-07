from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
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
                ("vibration_estimate", "Vibration Estimate", "Run estimate calculations with input parameters."),
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
