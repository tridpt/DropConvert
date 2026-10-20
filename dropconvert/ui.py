from __future__ import annotations

import getpass
import json
import sys
import unicodedata
from pathlib import Path

from pypdf import PdfReader
from PySide6.QtCore import QEasingCurve, QPropertyAnimation, QRect, Qt, QThread, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFont, QIcon, QLinearGradient, QPainter, QPen, QPixmap
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSystemTrayIcon,
    QTabBar,
    QVBoxLayout,
    QWidget,
)

from .core import (
    Capabilities,
    Action,
    available_actions,
    detect_capabilities,
    normalize_files,
    preview_text,
    run_action,
    unavailable_reason,
)
from . import explorer
from .organizer import PageOrganizer
from .queue import QueueDialog


STYLESHEET = """
QWidget { font-family: 'Segoe UI'; color: #e9eaf5; font-size: 12px; }
QDialog { background: #191b2c; }
QFrame#Card { background: #191b2c; border: 1px solid #343751; border-radius: 18px; }
QLabel#Title { font-size: 20px; font-weight: 700; color: #ffffff; }
QLabel#Muted { color: #a6a9c2; }
QLabel#Section { color: #bfc3dc; font-size: 11px; font-weight: 700; }
QLineEdit { background: #22243a; border: 1px solid #414560; border-radius: 9px;
            padding: 6px 9px; selection-background-color: #5a50ba; }
QLineEdit:focus { border-color: #897bff; }
QTabBar::tab { background: #282b43; border: 1px solid #414560; color: #bfc3dc;
               border-radius: 8px; padding: 6px 9px; margin-right: 4px; }
QTabBar::tab:hover { background: #393d5d; }
QTabBar::tab:selected { background: #6457e9; color: #ffffff; border-color: #8277ff; }
QPushButton { background: #2c3049; color: #f4f4fb; border: 1px solid #414560;
              border-radius: 10px; padding: 8px 12px; font-weight: 600; }
QPushButton:hover { background: #393d5d; border-color: #7770ed; }
QPushButton:pressed { background: #5147a7; }
QPushButton:disabled { color: #85889d; background: #242638; border-color: #34364a; }
QPushButton#Primary { background: #6457e9; border-color: #8277ff; }
QPushButton#Primary:hover { background: #786cff; }
QPushButton#Small { padding: 5px 9px; }
QFrame#ActionTile { background: #2c3049; border: 1px solid #414560; border-radius: 10px; }
QFrame#ActionTile:hover { background: #393d5d; border-color: #7770ed; }
QFrame#ActionTile:disabled { background: #242638; border-color: #34364a; }
QFrame#ActionTile QLabel#ActionTitle { color: #ffffff; font-weight: 700; }
QFrame#ActionTile QLabel#ActionDescription { color: #b6b9cd; font-size: 10px; }
QFrame#ActionTile:disabled QLabel { color: #85889d; }
QListWidget, QPlainTextEdit { background: #22243a; border: 1px solid #383b56;
                              border-radius: 9px; padding: 5px; selection-background-color: #5a50ba; }
QScrollArea { border: 0; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }
QProgressBar { background: #292c42; border: 0; border-radius: 4px; height: 7px; }
QProgressBar::chunk { background: #897bff; border-radius: 4px; }
QMenu { background: #24263a; border: 1px solid #474a68; padding: 5px; }
QMenu::item { padding: 8px 22px; }
QMenu::item:selected { background: #5d52cc; }
QScrollBar:vertical { background: #1d2032; width: 10px; margin: 0; border: 0; }
QScrollBar::handle:vertical { background: #535879; border-radius: 5px; min-height: 28px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
"""


def _paths_from_drop(event) -> tuple[Path, ...]:
    if not event.mimeData().hasUrls():
        return ()
    return normalize_files(url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile())


def _search_text(value: str) -> str:
    text = unicodedata.normalize("NFD", value.casefold()).replace("đ", "d")
    return "".join(character for character in text if unicodedata.category(character) != "Mn")


def _recommended_actions(files: tuple[Path, ...], actions: tuple[Action, ...], capabilities: Capabilities, scanned_pdf: bool = False) -> tuple[Action, ...]:
    if not files:
        return ()
    suffix = files[0].suffix.lower()
    if suffix == ".pdf":
        if len(files) > 1:
            keys = ("merge_pdf", "compress_pdf", "pdf_ocr_searchable", "pdf_docx")
        elif scanned_pdf:
            keys = ("pdf_ocr_searchable", "pdf_ocr_docx", "pdf_ocr_txt", "compress_pdf")
        else:
            keys = ("pdf_docx", "pdf_txt", "compress_pdf", "organize_pdf")
    elif suffix == ".docx":
        keys = ("docx_pdf", "docx_txt", "docx_md")
    elif suffix == ".doc":
        keys = ("doc_docx", "doc_pdf")
    elif suffix in {".txt", ".md"}:
        keys = ("text_docx", "text_pdf")
    else:
        keys = ("image_pdf", "image_ocr_searchable", "image_ocr_txt")
    by_key = {action.key: action for action in actions}
    available = tuple(by_key[key] for key in keys if key in by_key and not unavailable_reason(by_key[key], capabilities))
    return available or tuple(by_key[key] for key in keys if key in by_key)


class FloatingButton(QWidget):
    files_dropped = Signal(object)
    activated = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAcceptDrops(True)
        self.setToolTip("Thả tệp vào đây hoặc nhấp để mở Thả & Đổi")
        self._hover = False
        self._dragging = False
        self._base = QRect(0, 0, 76, 76)
        self._expanded = QRect(0, 0, 146, 76)
        self._animation = QPropertyAnimation(self, b"geometry", self)
        self._animation.setDuration(150)
        self._animation.setEasingCurve(QEasingCurve.OutCubic)
        self._tray_menu: QMenu | None = None

    def position_on_screen(self) -> None:
        screen = QApplication.primaryScreen().availableGeometry()
        x = screen.right() - 88
        y = screen.top() + round(screen.height() * 0.40)
        self._base = QRect(x, y, 76, 76)
        self._expanded = QRect(x - 70, y, 146, 76)
        self.setGeometry(self._base)

    def set_menu(self, menu: QMenu) -> None:
        self._tray_menu = menu

    def _expand(self, expand: bool) -> None:
        self._animation.stop()
        self._animation.setStartValue(self.geometry())
        self._animation.setEndValue(self._expanded if expand else self._base)
        self._animation.start()
        self._hover = expand
        self.update()

    def dragEnterEvent(self, event) -> None:
        if _paths_from_drop(event):
            event.acceptProposedAction()
            self._dragging = True
            self._expand(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:
        if _paths_from_drop(event):
            event.acceptProposedAction()

    def dragLeaveEvent(self, event) -> None:
        self._dragging = False
        self._expand(False)
        event.accept()

    def dropEvent(self, event) -> None:
        files = _paths_from_drop(event)
        self._dragging = False
        self._expand(False)
        if files:
            event.acceptProposedAction()
            self.files_dropped.emit(files)

    def enterEvent(self, event) -> None:
        if not self._dragging:
            self._hover = True
            self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        if not self._dragging:
            self._hover = False
            self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.LeftButton:
            self.activated.emit()
        elif event.button() == Qt.RightButton and self._tray_menu:
            self._tray_menu.exec(event.globalPosition().toPoint())
        super().mouseReleaseEvent(event)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        bounds = self.rect().adjusted(4, 4, -4, -4)
        gradient = QLinearGradient(bounds.topLeft(), bounds.bottomRight())
        gradient.setColorAt(0, QColor("#4036a4") if not self._hover else QColor("#6555ed"))
        gradient.setColorAt(1, QColor("#7669f5") if not self._hover else QColor("#978bff"))
        painter.setBrush(gradient)
        painter.setPen(QPen(QColor("#a69cff"), 1.4))
        painter.drawRoundedRect(bounds, 19, 19)
        # Document outline; drawn as paths so no icon font is required.
        left = 23 if self.width() < 100 else 16
        outline_pen = QPen(QColor("white"), 2.1)
        outline_pen.setCapStyle(Qt.RoundCap)
        outline_pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(outline_pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(left, 18, 28, 38, 3, 3)
        painter.drawLine(left + 6, 31, left + 21, 31)
        painter.drawLine(left + 6, 38, left + 21, 38)
        painter.drawLine(left + 6, 45, left + 17, 45)
        if self.width() >= 100:
            painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
            painter.drawText(57, 19, 78, 40, Qt.AlignVCenter | Qt.AlignLeft, "THẢ TỆP")
        else:
            painter.setFont(QFont("Segoe UI Symbol", 13, QFont.Bold))
            painter.drawText(40, 42, 18, 23, Qt.AlignCenter, "↔")


class ConversionWorker(QThread):
    progress_changed = Signal(int, str)
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, key: str, files: tuple[Path, ...], output_dir: Path, capabilities: Capabilities, password: str | None, compression: str = "balanced"):
        super().__init__()
        self.key = key
        self.files = files
        self.output_dir = output_dir
        self.capabilities = capabilities
        self.password = password
        self.compression = compression

    def run(self) -> None:
        try:
            result = run_action(
                self.key,
                self.files,
                self.output_dir,
                self.capabilities,
                self.password,
                lambda percent, message: self.progress_changed.emit(percent, message),
                self.compression,
            )
            self.succeeded.emit(result)
        except Exception as error:
            self.failed.emit(str(error) or error.__class__.__name__)


class ActionTile(QFrame):
    activated = Signal(str)

    def __init__(self, action: Action):
        super().__init__()
        self.action = action
        self.setObjectName("ActionTile")
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setAccessibleName(action.label)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(73)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 7)
        layout.setSpacing(1)
        title = QLabel(action.label)
        title.setObjectName("ActionTitle")
        title.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        layout.addWidget(title)
        description = QLabel(action.description)
        description.setObjectName("ActionDescription")
        description.setWordWrap(True)
        description.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        layout.addWidget(description)

    def mouseReleaseEvent(self, event) -> None:
        if self.isEnabled() and event.button() == Qt.LeftButton and self.rect().contains(event.position().toPoint()):
            self.activated.emit(self.action.key)
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event) -> None:
        if self.isEnabled() and event.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Space):
            self.activated.emit(self.action.key)
        else:
            super().keyPressEvent(event)


class FilePanel(QWidget):
    def __init__(self, capabilities: Capabilities):
        super().__init__()
        self.capabilities = capabilities
        self.files: tuple[Path, ...] = ()
        self.custom_output_dir: Path | None = None
        self.last_results: tuple[Path, ...] = ()
        self.worker: ConversionWorker | None = None
        self.queue_dialog: QueueDialog | None = None
        self.action_tiles: list[ActionTile] = []
        self.all_actions: tuple[Action, ...] = ()
        self.scanned_pdf = False
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAcceptDrops(True)
        self.resize(490, 710)
        self._build()
        self._refresh()

    def _build(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        card = QFrame()
        card.setObjectName("Card")
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 90))
        card.setGraphicsEffect(shadow)
        root.addWidget(card)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 17)
        layout.setSpacing(10)

        top = QHBoxLayout()
        title = QLabel("Thả & Đổi")
        title.setObjectName("Title")
        top.addWidget(title)
        top.addStretch()
        self.explorer_button = QPushButton()
        self.explorer_button.setObjectName("Small")
        self.explorer_button.setToolTip("Bật hoặc tắt lệnh 'Mở bằng Thả & Đổi' trong File Explorer")
        self.explorer_button.clicked.connect(self._toggle_explorer)
        self.explorer_button.setVisible(explorer.available())
        top.addWidget(self.explorer_button)
        self._refresh_explorer_button()
        close = QPushButton("✕")
        close.setObjectName("Small")
        close.setFixedSize(31, 30)
        close.setToolTip("Ẩn menu; nút thả tệp vẫn mở")
        close.clicked.connect(self.hide)
        top.addWidget(close)
        layout.addLayout(top)

        subtitle = QLabel("Kéo tệp vào nút nổi hoặc chọn tệp từ máy")
        subtitle.setObjectName("Muted")
        layout.addWidget(subtitle)

        choose = QPushButton("＋  Chọn tệp")
        choose.setObjectName("Primary")
        choose.clicked.connect(self._choose_files)
        layout.addWidget(choose)
        self.choose_button = choose

        files_head = QHBoxLayout()
        files_label = QLabel("TỆP ĐÃ CHỌN")
        files_label.setObjectName("Section")
        files_head.addWidget(files_label)
        files_head.addStretch()
        self.preview_button = QPushButton("Xem nhanh ▾")
        self.preview_button.setObjectName("Small")
        self.preview_button.setCheckable(True)
        self.preview_button.setToolTip("Hiện hoặc ẩn nội dung tệp đầu tiên")
        self.preview_button.toggled.connect(self._toggle_preview)
        files_head.addWidget(self.preview_button)
        self.count_label = QLabel("")
        self.count_label.setObjectName("Muted")
        files_head.addWidget(self.count_label)
        layout.addLayout(files_head)

        self.file_list = QListWidget()
        self.file_list.setFixedHeight(80)
        self.file_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.file_list.setToolTip("Kéo các dòng để đổi thứ tự trước khi gộp PDF")
        layout.addWidget(self.file_list)

        self.queue_button = QPushButton("☷  Mở hàng đợi từng tệp")
        self.queue_button.setToolTip("Chọn thao tác riêng cho từng tệp, kể cả nhiều định dạng khác nhau")
        self.queue_button.clicked.connect(self._open_queue)
        layout.addWidget(self.queue_button)

        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setFixedHeight(56)
        self.preview.hide()
        layout.addWidget(self.preview)

        self.action_label = QLabel("CHỌN CHỨC NĂNG")
        self.action_label.setObjectName("Section")
        layout.addWidget(self.action_label)
        self.group_tabs = QTabBar()
        self.group_tabs.setDrawBase(False)
        self.group_tabs.setExpanding(False)
        self.group_tabs.currentChanged.connect(self._render_actions)
        layout.addWidget(self.group_tabs)
        self.action_search = QLineEdit()
        self.action_search.setPlaceholderText("Tìm chức năng, ví dụ: OCR, nén, Word…")
        self.action_search.setClearButtonEnabled(True)
        self.action_search.textChanged.connect(self._render_actions)
        layout.addWidget(self.action_search)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setMinimumHeight(160)
        self.action_host = QWidget()
        self.action_layout = QGridLayout(self.action_host)
        self.action_layout.setContentsMargins(0, 0, 5, 0)
        self.action_layout.setHorizontalSpacing(8)
        self.action_layout.setVerticalSpacing(8)
        self.action_layout.setColumnStretch(0, 1)
        self.action_layout.setColumnStretch(1, 1)
        self.scroll.setWidget(self.action_host)
        layout.addWidget(self.scroll, 1)

        output_row = QHBoxLayout()
        output_info = QVBoxLayout()
        output_title = QLabel("NƠI LƯU")
        output_title.setObjectName("Section")
        output_info.addWidget(output_title)
        self.output_label = QLabel("")
        self.output_label.setObjectName("Muted")
        self.output_label.setMaximumWidth(290)
        output_info.addWidget(self.output_label)
        output_row.addLayout(output_info, 1)
        output_change = QPushButton("Đổi…")
        output_change.setObjectName("Small")
        output_change.clicked.connect(self._choose_output)
        output_row.addWidget(output_change)
        self.output_change = output_change
        layout.addLayout(output_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        self.status = QLabel("")
        self.status.setObjectName("Muted")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        bottom = QHBoxLayout()
        bottom.addStretch()
        self.open_file = QPushButton("Mở tệp")
        self.open_file.setObjectName("Small")
        self.open_file.clicked.connect(self._open_result)
        self.open_file.hide()
        bottom.addWidget(self.open_file)
        self.open_folder = QPushButton("Mở thư mục")
        self.open_folder.setObjectName("Small")
        self.open_folder.clicked.connect(self._open_folder)
        self.open_folder.hide()
        bottom.addWidget(self.open_folder)
        layout.addLayout(bottom)


    def place_near(self, button: FloatingButton) -> None:
        screen = QApplication.screenAt(button.geometry().center()) or QApplication.primaryScreen()
        available = screen.availableGeometry()
        height = min(710, available.height() - 22)
        self.resize(490, height)
        x = button.x() - self.width() - 12
        if x < available.left():
            x = button.x() + button.width() + 12
        x = max(available.left() + 5, min(x, available.right() - self.width() - 5))
        y = max(available.top() + 5, min(button.y() - 140, available.bottom() - self.height() - 5))
        self.move(x, y)

    def dragEnterEvent(self, event) -> None:
        if _paths_from_drop(event):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:
        files = _paths_from_drop(event)
        if files:
            event.acceptProposedAction()
            self.set_files(files)

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key_Escape:
            self.hide()
        else:
            super().keyPressEvent(event)

    def set_files(self, files: tuple[Path, ...]) -> None:
        if self.worker and self.worker.isRunning():
            self.status.setText("Đang xử lý. Hãy đợi thao tác hiện tại hoàn tất.")
            return
        self.files = normalize_files(files)
        self.custom_output_dir = None
        self.last_results = ()
        self.file_list.clear()
        for file in self.files:
            item = QListWidgetItem(file.name)
            item.setData(Qt.UserRole, str(file))
            item.setToolTip(str(file))
            self.file_list.addItem(item)
        self._refresh()

    def _ordered_files(self) -> tuple[Path, ...]:
        return tuple(Path(self.file_list.item(index).data(Qt.UserRole)) for index in range(self.file_list.count()))

    def _default_output_dir(self) -> Path | None:
        files = self._ordered_files()
        return files[0].parent / "DaChuyenDoi" if files else None

    def _output_dir(self) -> Path | None:
        return self.custom_output_dir or self._default_output_dir()

    def _refresh(self) -> None:
        files = self._ordered_files()
        self.count_label.setText(f"{len(files)} tệp" if files else "Chưa có tệp")
        self.preview_button.setEnabled(bool(files))
        if not files and self.preview_button.isChecked():
            self.preview_button.setChecked(False)
        if self.preview_button.isChecked():
            self.preview.setPlainText(preview_text(files[0]) if files else "")
        folder = self._output_dir()
        self.output_label.setText(str(folder) if folder else "Cạnh tệp gốc / DaChuyenDoi")
        self.output_label.setToolTip(str(folder) if folder else "")
        self.all_actions = available_actions(files)
        self.scanned_pdf = False
        if len(files) == 1 and files[0].suffix.lower() == ".pdf":
            try:
                reader = PdfReader(str(files[0]))
                if reader.pages:
                    text = reader.pages[0].extract_text() or ""
                    self.scanned_pdf = sum(character.isalnum() for character in text) < 16
            except Exception:
                pass
        self.action_search.blockSignals(True)
        self.action_search.clear()
        self.action_search.blockSignals(False)
        self.group_tabs.blockSignals(True)
        while self.group_tabs.count():
            self.group_tabs.removeTab(0)
        if self.all_actions:
            groups = (("Gợi ý", "recommended"), ("Tất cả", "all"),
                      ("Chuyển đổi", "Chuyển đổi"), ("OCR", "Nhận dạng chữ"),
                      ("Công cụ PDF", "Công cụ PDF"))
            for label, key in groups:
                if key in {"recommended", "all"} or any(action.category == key for action in self.all_actions):
                    index = self.group_tabs.addTab(label)
                    self.group_tabs.setTabData(index, key)
            self.group_tabs.setCurrentIndex(0)
        self.group_tabs.blockSignals(False)
        self.group_tabs.setVisible(bool(self.all_actions))
        self.action_search.setVisible(bool(self.all_actions))
        self._render_actions()
        self.queue_button.setVisible(len(files) > 1)
        self.queue_button.setEnabled(len(files) > 1 and not (self.worker and self.worker.isRunning()))
        self.status.setText("Thứ tự trong danh sách sẽ dùng khi gộp PDF." if len(files) > 1 and all(p.suffix.lower() == ".pdf" for p in files) else "Chọn hàng đợi để xử lý tệp lẫn định dạng; tệp gốc luôn được giữ nguyên.")
        self.progress_bar.hide()
        self.open_file.hide()
        self.open_folder.hide()

    def _toggle_preview(self, expanded: bool) -> None:
        self.preview.setVisible(expanded)
        self.preview_button.setText("Ẩn xem nhanh ▴" if expanded else "Xem nhanh ▾")
        if expanded:
            files = self._ordered_files()
            self.preview.setPlainText(preview_text(files[0]) if files else "")

    def _render_actions(self, *_args) -> None:
        while self.action_layout.count():
            item = self.action_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.action_tiles.clear()
        files = self._ordered_files()
        query = _search_text(self.action_search.text().strip())
        if query:
            actions = tuple(action for action in self.all_actions
                            if query in _search_text(action.label + " " + action.description + " " + action.category))
            actions = tuple(sorted(actions, key=lambda action: (
                query not in _search_text(action.label),
                not _search_text(action.label).startswith(query),
            )))
            self.action_label.setText(f"KẾT QUẢ TÌM KIẾM · {len(actions)}")
        else:
            group = self.group_tabs.tabData(self.group_tabs.currentIndex()) if self.group_tabs.count() else None
            if group == "recommended":
                actions = _recommended_actions(files, self.all_actions, self.capabilities, self.scanned_pdf)
            elif group == "all":
                actions = self.all_actions
            else:
                actions = tuple(action for action in self.all_actions if action.category == group)
            self.action_label.setText(f"CHỌN CHỨC NĂNG · {len(actions)}" if actions else "CHỌN CHỨC NĂNG")
        if not actions:
            if query:
                message = "Không tìm thấy chức năng. Thử từ khóa khác."
            elif files:
                message = "Các tệp khác định dạng. Bấm ‘Mở hàng đợi từng tệp’ để chọn chức năng riêng."
            else:
                message = "Chọn hoặc thả tệp để xem các chức năng phù hợp."
            help_label = QLabel(message)
            help_label.setObjectName("Muted")
            help_label.setWordWrap(True)
            self.action_layout.addWidget(help_label, 0, 0, 1, 2)
        for row in range(20):
            self.action_layout.setRowStretch(row, 0)
        for index, action in enumerate(actions):
            tile = ActionTile(action)
            reason = unavailable_reason(action, self.capabilities)
            if reason:
                tile.setEnabled(False)
                tile.setToolTip(reason + "\n" + action.description)
            else:
                tile.setToolTip(action.description)
            tile.activated.connect(self._start)
            if self.worker and self.worker.isRunning():
                tile.setEnabled(False)
            self.action_tiles.append(tile)
            self.action_layout.addWidget(tile, index // 2, index % 2)
        self.action_layout.setRowStretch((len(actions) + 1) // 2, 1)
        self.scroll.verticalScrollBar().setValue(0)

    def _choose_files(self) -> None:
        filter_text = "Tài liệu và ảnh (*.docx *.doc *.pdf *.txt *.md *.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff)"
        paths, _ = QFileDialog.getOpenFileNames(self, "Chọn tệp", "", filter_text)
        if paths:
            self.set_files(tuple(Path(path) for path in paths))

    def _choose_output(self) -> None:
        initial = str(self._output_dir() or Path.home())
        path = QFileDialog.getExistingDirectory(self, "Chọn nơi lưu", initial)
        if path:
            self.custom_output_dir = Path(path)
            self.output_label.setText(path)
            self.output_label.setToolTip(path)

    def _open_queue(self) -> None:
        files = self._ordered_files()
        folder = self._output_dir()
        if not files or folder is None or (self.worker and self.worker.isRunning()):
            return
        if self.queue_dialog is None or not self.queue_dialog.isVisible():
            self.queue_dialog = QueueDialog(files, folder, self.capabilities, self)
            self.queue_dialog.show()
        else:
            self.queue_dialog.add_files(files)
            self.queue_dialog.raise_()
            self.queue_dialog.activateWindow()

    def _start(self, key: str) -> None:
        if self.worker and self.worker.isRunning():
            return
        files = self._ordered_files()
        folder = self._output_dir()
        if not files or folder is None:
            return
        action: Action | None = next((item for item in available_actions(files) if item.key == key), None)
        if action is None:
            return
        if key == "organize_pdf":
            dialog = PageOrganizer(files[0], folder, self)
            dialog.exec()
            if dialog.saved_path:
                from .core import Result

                self._succeeded(Result((dialog.saved_path,)))
            return
        password = None
        if action.requires_password:
            password, accepted = QInputDialog.getText(self, "Đặt mật khẩu PDF", "Mật khẩu để mở tệp:", QLineEdit.Password)
            if not accepted:
                return
            if not password:
                self.status.setText("Hãy nhập mật khẩu khác rỗng.")
                return
        compression = "balanced"
        if key == "compress_pdf":
            options = ("Rõ nét · 220 DPI", "Cân bằng · 160 DPI", "Dung lượng nhỏ · 110 DPI")
            selected, accepted = QInputDialog.getItem(self, "Nén PDF", "Chọn mức chất lượng:", options, 1, False)
            if not accepted:
                return
            compression = {options[0]: "clear", options[1]: "balanced", options[2]: "small"}[selected]
        self.last_results = ()
        self.worker = ConversionWorker(key, files, folder, self.capabilities, password, compression)
        self.worker.progress_changed.connect(self._update_progress)
        self.worker.succeeded.connect(self._succeeded)
        self.worker.failed.connect(self._failed)
        self.worker.finished.connect(self._worker_finished)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.status.setText("Đang bắt đầu…")
        self.choose_button.setEnabled(False)
        self.output_change.setEnabled(False)
        self.queue_button.setEnabled(False)
        self.group_tabs.setEnabled(False)
        self.action_search.setEnabled(False)
        for tile in self.action_tiles:
            tile.setEnabled(False)
        self.worker.start()

    def _update_progress(self, percent: int, message: str) -> None:
        self.progress_bar.setValue(percent)
        self.status.setText(message)

    def _succeeded(self, result) -> None:
        self.last_results = result.files
        self.progress_bar.setValue(100)
        message = f"Xong: {len(result.files)} tệp đã lưu."
        if result.note:
            message += " " + result.note
        self.status.setText(message)
        self.open_folder.show()
        if len(result.files) == 1:
            self.open_file.show()

    def _failed(self, message: str) -> None:
        self.progress_bar.hide()
        self.status.setText("Không xử lý được: " + message)

    def _worker_finished(self) -> None:
        self.choose_button.setEnabled(True)
        self.output_change.setEnabled(True)
        self.queue_button.setEnabled(True)
        self.group_tabs.setEnabled(True)
        self.action_search.setEnabled(True)
        for tile in self.action_tiles:
            tile.setEnabled(not unavailable_reason(tile.action, self.capabilities))

    def _open_result(self) -> None:
        if self.last_results:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_results[0])))

    def _open_folder(self) -> None:
        if self.last_results:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_results[0].parent)))

    def _refresh_explorer_button(self) -> None:
        if explorer.available():
            self.explorer_button.setText(
                "Tắt chuột phải" if explorer.is_registered() else "Bật chuột phải"
            )

    def _toggle_explorer(self) -> None:
        try:
            if explorer.is_registered():
                explorer.unregister()
                self.status.setText("Đã tắt lệnh chuột phải trong File Explorer.")
            else:
                explorer.register()
                self.status.setText("Đã bật chuột phải và Gửi tới. Chọn nhiều tệp rồi dùng Gửi tới → Thả & Đổi.")
        except Exception as error:
            self.status.setText("Không cập nhật được File Explorer: " + str(error))
        self._refresh_explorer_button()


def make_icon() -> QIcon:
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QColor("#6c5ce7"))
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(2, 2, 60, 60, 15, 15)
    painter.setPen(QPen(QColor("white"), 3))
    painter.setBrush(Qt.NoBrush)
    painter.drawRoundedRect(18, 12, 29, 40, 3, 3)
    painter.drawLine(24, 27, 40, 27)
    painter.drawLine(24, 35, 40, 35)
    painter.end()
    return QIcon(pixmap)


def main() -> int:
    app = QApplication([])
    app.setApplicationName("Thả & Đổi")
    app.setApplicationVersion("0.5.0")
    app.setQuitOnLastWindowClosed(False)
    app.setStyleSheet(STYLESHEET)
    server_name = "ThaVaDoi-" + getpass.getuser()
    initial_files = normalize_files(sys.argv[2:] if len(sys.argv) > 1 and sys.argv[1] == "--open" else ())
    existing = QLocalSocket()
    existing.connectToServer(server_name)
    if existing.waitForConnected(250):
        existing.write((json.dumps({"files": [str(path) for path in initial_files]}, ensure_ascii=False) + "\n").encode("utf-8"))
        existing.waitForBytesWritten(1000)
        existing.disconnectFromServer()
        return 0
    server = QLocalServer(app)
    if not server.listen(server_name):
        QLocalServer.removeServer(server_name)
        server.listen(server_name)
    icon = make_icon()
    app.setWindowIcon(icon)
    capabilities = detect_capabilities()
    button = FloatingButton()
    button.position_on_screen()
    panel = FilePanel(capabilities)
    if initial_files:
        panel.set_files(initial_files)

    def show_panel() -> None:
        panel.place_near(button)
        panel.show()
        panel.raise_()
        panel.activateWindow()

    def on_new_connection() -> None:
        while server.hasPendingConnections():
            connection = server.nextPendingConnection()
            buffer = bytearray()

            def receive(client=connection, data=buffer) -> None:
                data.extend(bytes(client.readAll()))
                while b"\n" in data:
                    line, _, rest = data.partition(b"\n")
                    data[:] = rest
                    try:
                        requested = normalize_files(json.loads(line.decode("utf-8")).get("files", ()))
                    except (ValueError, TypeError, AttributeError):
                        requested = ()
                    if requested:
                        panel.set_files(requested)
                    show_panel()

            connection.readyRead.connect(receive)
            if connection.bytesAvailable():
                receive()
            connection.disconnected.connect(connection.deleteLater)

    server.newConnection.connect(on_new_connection)

    button.activated.connect(show_panel)
    button.files_dropped.connect(lambda files: (panel.set_files(files), show_panel()))
    menu = QMenu()
    menu.addAction("Mở Thả & Đổi", show_panel)
    if explorer.available():
        context_action = menu.addAction("")
        context_action.triggered.connect(panel._toggle_explorer)
        menu.aboutToShow.connect(lambda: context_action.setText(
            "Tắt lệnh chuột phải File Explorer" if explorer.is_registered() else "Bật lệnh chuột phải File Explorer"
        ))
    menu.addAction("Hiện / ẩn nút nổi", lambda: button.hide() if button.isVisible() else button.show())
    menu.addSeparator()
    menu.addAction("Thoát", app.quit)
    button.set_menu(menu)
    tray = QSystemTrayIcon(icon, app)
    tray.setToolTip("Thả & Đổi — chuyển đổi tệp trên máy")
    tray.setContextMenu(menu)
    tray.activated.connect(lambda reason: show_panel() if reason == QSystemTrayIcon.DoubleClick else None)
    if QSystemTrayIcon.isSystemTrayAvailable():
        tray.show()
    button.show()
    show_panel()
    return app.exec()
