"""A sequential conversion queue for mixed document types."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from threading import Event

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDialog, QFileDialog, QHBoxLayout,
    QInputDialog, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from .core import ACTIONS, COMPRESSION_PRESETS, Capabilities, available_actions, normalize_files, run_action, unavailable_reason


@dataclass(frozen=True)
class QueueTask:
    row: int
    source: Path
    action: str
    compression: str = "balanced"
    password: str | None = None


class QueueWorker(QThread):
    progress_changed = Signal(int, int, str)
    task_succeeded = Signal(int, object)
    task_failed = Signal(int, str)

    def __init__(self, tasks: tuple[QueueTask, ...], output_dir: Path, capabilities: Capabilities):
        super().__init__()
        self.tasks = tasks
        self.output_dir = output_dir
        self.capabilities = capabilities
        self._stop = Event()

    def request_stop(self) -> None:
        self._stop.set()

    def run(self) -> None:
        for task in self.tasks:
            if self._stop.is_set():
                break
            try:
                result = run_action(
                    task.action, (task.source,), self.output_dir, self.capabilities,
                    task.password,
                    lambda percent, message, row=task.row: self.progress_changed.emit(row, percent, message),
                    task.compression,
                )
                self.task_succeeded.emit(task.row, result)
            except Exception as error:
                self.task_failed.emit(task.row, str(error) or error.__class__.__name__)


class QueueDialog(QDialog):
    def __init__(self, files: tuple[Path, ...], output_dir: Path, capabilities: Capabilities, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Hàng đợi — Thả & Đổi")
        self.setMinimumSize(840, 430)
        self.setAcceptDrops(True)
        self.capabilities = capabilities
        self.output_dir = output_dir
        self.worker: QueueWorker | None = None
        self.results: dict[int, tuple[Path, ...]] = {}

        root = QVBoxLayout(self)
        root.addWidget(QLabel("Mỗi tệp chọn một thao tác. Hàng đợi chạy lần lượt và giữ tệp gốc."))
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(("Tệp", "Thao tác", "Mức nén PDF", "Tiến độ / trạng thái"))
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 245)
        self.table.setColumnWidth(1, 195)
        self.table.setColumnWidth(2, 130)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        root.addWidget(self.table)

        self.summary = QLabel("")
        root.addWidget(self.summary)
        buttons = QHBoxLayout()
        self.add_button = QPushButton("＋ Thêm tệp")
        self.add_button.clicked.connect(self._choose_files)
        buttons.addWidget(self.add_button)
        self.output_button = QPushButton("Nơi lưu…")
        self.output_button.clicked.connect(self._choose_output)
        buttons.addWidget(self.output_button)
        buttons.addStretch()
        self.open_button = QPushButton("Mở thư mục")
        self.open_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_dir))))
        self.open_button.setEnabled(False)
        buttons.addWidget(self.open_button)
        self.retry_button = QPushButton("Thử lại lỗi")
        self.retry_button.clicked.connect(lambda: self._start(retry_only=True))
        self.retry_button.setEnabled(False)
        buttons.addWidget(self.retry_button)
        self.stop_button = QPushButton("Dừng sau tệp này")
        self.stop_button.clicked.connect(self._stop)
        self.stop_button.setEnabled(False)
        buttons.addWidget(self.stop_button)
        self.start_button = QPushButton("Bắt đầu")
        self.start_button.setObjectName("Primary")
        self.start_button.clicked.connect(self._start)
        buttons.addWidget(self.start_button)
        root.addLayout(buttons)
        self.add_files(files)
        QApplication.instance().aboutToQuit.connect(self._shutdown)

    def add_files(self, files: tuple[Path, ...]) -> None:
        known = {Path(self.table.item(row, 0).data(Qt.UserRole)) for row in range(self.table.rowCount())}
        for source in normalize_files(files):
            if source in known:
                continue
            known.add(source)
            actions = [action for action in available_actions((source,))
                       if action.key not in {"organize_pdf", "merge_pdf"}]
            row = self.table.rowCount()
            self.table.insertRow(row)
            item = QTableWidgetItem(source.name)
            item.setData(Qt.UserRole, str(source))
            item.setToolTip(str(source))
            self.table.setItem(row, 0, item)
            combo = QComboBox()
            for action in actions:
                if not unavailable_reason(action, self.capabilities):
                    combo.addItem(action.label, action.key)
            if combo.count() == 0:
                combo.addItem("Không có thao tác khả dụng", "")
                combo.setEnabled(False)
            preferred = {".docx": "docx_pdf", ".doc": "doc_docx", ".pdf": "compress_pdf"}.get(source.suffix.lower())
            if preferred:
                index = combo.findData(preferred)
                if index >= 0:
                    combo.setCurrentIndex(index)
            combo.currentIndexChanged.connect(lambda _index, r=row: self._action_changed(r))
            self.table.setCellWidget(row, 1, combo)
            preset = QComboBox()
            for key, (_, _, label) in COMPRESSION_PRESETS.items():
                preset.addItem(label.capitalize(), key)
            preset.setCurrentIndex(preset.findData("balanced"))
            self.table.setCellWidget(row, 2, preset)
            self._action_changed(row)
            status = QTableWidgetItem("Chờ")
            status.setData(Qt.UserRole, "pending")
            self.table.setItem(row, 3, status)
        self._update_summary()

    def _action_changed(self, row: int) -> None:
        action = self.table.cellWidget(row, 1).currentData()
        self.table.cellWidget(row, 2).setEnabled(action == "compress_pdf")
        status = self.table.item(row, 3)
        if status:
            status.setText("Chờ")
            status.setData(Qt.UserRole, "pending")
            self.results.pop(row, None)

    def _choose_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Thêm tệp vào hàng đợi")
        if paths:
            self.add_files(tuple(Path(path) for path in paths))

    def _choose_output(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Chọn nơi lưu", str(self.output_dir))
        if selected:
            self.output_dir = Path(selected)
            self._update_summary()

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls() and any(url.isLocalFile() for url in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:
        files = normalize_files(url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile())
        if files and not (self.worker and self.worker.isRunning()):
            self.add_files(files)
            event.acceptProposedAction()

    def _start(self, retry_only: bool = False) -> None:
        if self.worker and self.worker.isRunning():
            return
        tasks = []
        for row in range(self.table.rowCount()):
            status = self.table.item(row, 3)
            state = status.data(Qt.UserRole)
            if retry_only and state != "failed":
                continue
            if not retry_only and state == "done":
                continue
            key = self.table.cellWidget(row, 1).currentData()
            if not key:
                continue
            password = None
            if ACTIONS[key].requires_password:
                password, accepted = QInputDialog.getText(self, "Đặt mật khẩu PDF", f"Mật khẩu cho {self.table.item(row, 0).text()}:", QLineEdit.Password)
                if not accepted:
                    return
                if not password:
                    status.setText("Cần nhập mật khẩu")
                    status.setData(Qt.UserRole, "failed")
                    continue
            source = Path(self.table.item(row, 0).data(Qt.UserRole))
            compression = self.table.cellWidget(row, 2).currentData()
            tasks.append(QueueTask(row, source, key, compression, password))
            status.setText("Chờ")
            status.setData(Qt.UserRole, "pending")
        if not tasks:
            self._update_summary()
            return
        self.worker = QueueWorker(tuple(tasks), self.output_dir, self.capabilities)
        self.worker.progress_changed.connect(self._progress)
        self.worker.task_succeeded.connect(self._succeeded)
        self.worker.task_failed.connect(self._failed)
        self.worker.finished.connect(self._finished)
        self.add_button.setEnabled(False)
        self.output_button.setEnabled(False)
        self.start_button.setEnabled(False)
        self.retry_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        for row in range(self.table.rowCount()):
            self.table.cellWidget(row, 1).setEnabled(False)
            self.table.cellWidget(row, 2).setEnabled(False)
        self.worker.start()

    def _progress(self, row: int, percent: int, message: str) -> None:
        status = self.table.item(row, 3)
        status.setText(f"{percent}% · {message}")
        status.setData(Qt.UserRole, "running")

    def _succeeded(self, row: int, result) -> None:
        self.results[row] = result.files
        status = self.table.item(row, 3)
        status.setText(f"Xong · {len(result.files)} tệp" + (f" · {result.note}" if result.note else ""))
        status.setToolTip("\n".join(str(path) for path in result.files))
        status.setData(Qt.UserRole, "done")
        self.open_button.setEnabled(True)
        self._update_summary()

    def _failed(self, row: int, message: str) -> None:
        status = self.table.item(row, 3)
        status.setText("Lỗi · " + message)
        status.setToolTip(message)
        status.setData(Qt.UserRole, "failed")
        self._update_summary()

    def _stop(self) -> None:
        if self.worker:
            self.worker.request_stop()
            self.stop_button.setEnabled(False)
            self.stop_button.setText("Đang dừng…")

    def _finished(self) -> None:
        self.add_button.setEnabled(True)
        self.output_button.setEnabled(True)
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.stop_button.setText("Dừng sau tệp này")
        for row in range(self.table.rowCount()):
            combo = self.table.cellWidget(row, 1)
            combo.setEnabled(bool(combo.currentData()))
            self._action_changed_enabled(row)
        self._update_summary()

    def _action_changed_enabled(self, row: int) -> None:
        self.table.cellWidget(row, 2).setEnabled(self.table.cellWidget(row, 1).currentData() == "compress_pdf")

    def _update_summary(self) -> None:
        states = [self.table.item(row, 3).data(Qt.UserRole) for row in range(self.table.rowCount()) if self.table.item(row, 3)]
        done = states.count("done")
        failed = states.count("failed")
        self.summary.setText(f"{done}/{self.table.rowCount()} xong · {failed} lỗi · Lưu tại {self.output_dir}")
        self.summary.setToolTip(str(self.output_dir))
        self.retry_button.setEnabled(failed > 0 and not (self.worker and self.worker.isRunning()))

    def closeEvent(self, event) -> None:
        if self.worker and self.worker.isRunning():
            event.ignore()
            self._stop()
        else:
            super().closeEvent(event)

    def _shutdown(self) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.worker.wait()
