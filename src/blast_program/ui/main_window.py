from PySide6.QtWidgets import QMainWindow, QStackedWidget

from blast_program.ui.screens import PlaceholderScreen, StartScreen


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Blast Program")
        self.resize(900, 620)

        self._stack = QStackedWidget()
        self._stack.setObjectName("mainStack")
        self.setCentralWidget(self._stack)
        self._screens = {}
        self._apply_theme()
        self._build_screens()
        self.navigate_to("start")

    def _apply_theme(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #1f4fa3;
            }
            QStackedWidget#mainStack {
                background-color: #1f4fa3;
            }
            QLabel#pageTitle {
                color: #1a202c;
                font-size: 29px;
                font-weight: 700;
                letter-spacing: 0.2px;
            }
            QLabel#pageSubtitle {
                color: #4a5568;
                font-size: 14px;
                line-height: 1.35em;
            }
            QLabel#bodyText {
                color: #2d3748;
                font-size: 14px;
                line-height: 1.35em;
            }
            QWidget#contentPanel {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            QPushButton {
                color: #f8fafc;
                background-color: #1f6feb;
                border: 1px solid #1b5ec9;
                border-radius: 10px;
                padding: 10px 14px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #1b5ec9;
            }
            QPushButton:pressed {
                background-color: #194fa8;
            }
            QPushButton:focus {
                border: 2px solid #0a4db5;
            }
            QPushButton[variant="secondary"] {
                color: #1f2937;
                background-color: #ffffff;
                border: 1px solid #cbd5e0;
            }
            QPushButton[variant="secondary"]:hover {
                background-color: #f7fafc;
            }
            QPushButton[variant="card"] {
                color: #1a202c;
                background-color: #d1d5db;
                border: 1px solid #9ca3af;
                border-radius: 12px;
                text-align: center;
                padding: 14px 16px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton[variant="card"]:hover {
                border: 1px solid #6b7280;
                background-color: #c4c9d1;
            }
            QPushButton[variant="card"]:pressed {
                background-color: #b8bec8;
            }
            """
        )

    def _build_screens(self) -> None:
        self._register_screen("start", StartScreen(self.navigate_to))
        self._register_screen(
            "vibration_estimate",
            PlaceholderScreen(
                "Vibration Estimate",
                "Section scaffolded. We will implement formulas, inputs, and results in a later section.",
                self.navigate_home,
            ),
        )
        self._register_screen(
            "timing_solver",
            PlaceholderScreen(
                "Timing Solver",
                "Section scaffolded. We will implement timing calculation logic in a later section.",
                self.navigate_home,
            ),
        )
        self._register_screen(
            "empirical_formula",
            PlaceholderScreen(
                "Empirical Formula Calculation",
                "Section scaffolded. We will add chemistry inputs and solver logic in a later section.",
                self.navigate_home,
            ),
        )
        self._register_screen(
            "diagram_maker",
            PlaceholderScreen(
                "Diagram Maker",
                "Section scaffolded. We will add drawing/export features and integrate Timing Solver here later.",
                self.navigate_home,
            ),
        )
        self._register_screen(
            "quick_cheat_sheets",
            PlaceholderScreen(
                "Small Quick Cheat Sheets",
                "Section scaffolded. We will connect this to Cheat Sheets / Knowledge content in the next section.",
                self.navigate_home,
            ),
        )

    def _register_screen(self, key: str, widget) -> None:
        self._screens[key] = widget
        self._stack.addWidget(widget)

    def navigate_home(self) -> None:
        self.navigate_to("start")

    def navigate_to(self, key: str) -> None:
        widget = self._screens.get(key)
        if widget is not None:
            self._stack.setCurrentWidget(widget)
