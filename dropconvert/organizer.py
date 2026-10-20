"""Visual PDF page organizer; original PDFs are never modified."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QThread, Signal
from PySide6.QtGui import QIcon, QPixmap, QTransform
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
)

from .core import rearrange_pdf


class ThumbnailWorker(QThread):
    page_ready = Signal(int, bytes)
    failed = Signal(str)

    def __init__(self, source: Path):
        super().__init__()
        self.source = source

    def run(self) -> None:
        try:
            import pypdfium2 as pdfium

            pdf = pdfium.PdfDocument(str(self.source))
            try:
                for index in range(len(pdf)):
                    if self.isInterruptionRequested():
                        break
                    page = pdf[index]
                    try:
                        image = page.render(scale=0.55).to_pil()
                        try:
                            image.thumbnail((118, 145))
                            with BytesIO() as buffer:
                                image.save(buffer, "PNG")
                                self.page_ready.emit(index, buffer.getvalue())
                        finally:
                            image.close()
                    finally:
                        page.close()
            finally:
                pdf.close()
        except Exception as error:
            self.failed.emit(str(error) or error.__class__.__name__)


class SaveWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, source: Path, pages: tuple[tuple[int, int], ...], output_dir: Path):
        super().__init__()
        self.source = source
        self.pages = pages
        self.output_dir = output_dir

    def run(self) -> None:
        try:
            self.succeeded.emit(rearrange_pdf(self.source, self.pages, self.output_dir))
        except Exception as error:
            self.failed.emit(str(error) or error.__class__.__name__)


class PageOrganizer(QDialog):
    def __init__(self, source: Path, output_dir: Path, parent=None):
        super().__init__(parent)
        self.source = source
        self.output_dir = output_dir
        self.saved_path: Path | None = None
        self._pixmaps: dict[int, QPixmap] = {}
        self._load_error = False
        self._save_worker: SaveWorker | None = None
        self.setWindowTitle("Sắp xếp trang PDF — Thả & Đổi")
        self.setMinimumSize(590, 590)
        self.resize(680, 710)
        self._build()
        self._thumbnail_worker = ThumbnailWorker(source)
        self._thumbnail_worker.page_ready.connect(self._add_page)
        self._thumbnail_worker.failed.connect(self._load_failed)
        self._thumbnail_worker.finished.connect(self._loading_finished)
        self._thumbnail_worker.start()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)
        title = QLabel("SẮP XẾP TRANG PDF")
        title.setObjectName("Title")
        layout.addWidget(title)
        description = QLabel("Kéo ảnh trang để đổi thứ tự. Chọn trang để xoay hoặc bỏ. Tệp gốc được giữ nguyên.")
        description.setWordWrap(True)
        description.setObjectName("Muted")
        layout.addWidget(description)
        self.pages = QListWidget()
        self.pages.setIconSize(QSize(118, 145))
        self.pages.setSpacing(5)
        self.pages.setDragDropMode(QAbstractItemView.InternalMove)
        self.pages.setDefaultDropAction(Qt.MoveAction)
        self.pages.setSelectionMode(QAbstractItemView.SingleSelection)
        self.pages.currentItemChanged.connect(self._selection_changed)
        layout.addWidget(self.pages, 1)
        tools = QHBoxLayout()
        self.left = QPushButton("↶ Xoay trái")
        self.right = QPushButton("↷ Xoay phải")
        self.remove = QPushButton("Bỏ trang")
        self.left.clicked.connect(lambda: self._rotate(-90))
        self.right.clicked.connect(lambda: self._rotate(90))
        self.remove.clicked.connect(self._remove)
        for widget in (self.left, self.right, self.remove):
            widget.setEnabled(False)
            tools.addWidget(widget)
        layout.addLayout(tools)
        self.status = QLabel("Đang tạo ảnh thu nhỏ…")
        self.status.setObjectName("Muted")
        layout.addWidget(self.status)
        bottom = QHBoxLayout()
        bottom.addStretch()
        cancel = QPushButton("Đóng")
        cancel.clicked.connect(self.reject)
        bottom.addWidget(cancel)
        self.save = QPushButton("Lưu PDF mới")
        self.save.setObjectName("Primary")
        self.save.setEnabled(False)
        self.save.clicked.connect(self._save)
        bottom.addWidget(self.save)
        layout.addLayout(bottom)

    def _add_page(self, index: int, data: bytes) -> None:
        pixmap = QPixmap()
        if not pixmap.loadFromData(data, "PNG"):
            self._load_failed("Không đọc được ảnh trang")
            return
        self._pixmaps[index] = pixmap
        item = QListWidgetItem(QIcon(pixmap), f"Trang {index + 1}")
        item.setData(Qt.UserRole, index)
        item.setData(Qt.UserRole + 1, 0)
        item.setSizeHint(QSize(350, 153))
        self.pages.addItem(item)
        if self.pages.currentItem() is None:
            self.pages.setCurrentItem(item)
        self.status.setText(f"Đã tải {self.pages.count()} trang…")

    def _load_failed(self, message: str) -> None:
        self._load_error = True
        self.status.setText("Không mở được PDF: " + message)

    def _loading_finished(self) -> None:
        if not self._load_error:
            self.status.setText(f"{self.pages.count()} trang. Kéo một trang để đổi vị trí.")
            self.save.setEnabled(self.pages.count() > 0)

    def _selection_changed(self, current, previous) -> None:
        enabled = current is not None and self._save_worker is None
        for widget in (self.left, self.right, self.remove):
            widget.setEnabled(enabled)

    def _rotate(self, degrees: int) -> None:
        item = self.pages.currentItem()
        if item is None:
            return
        rotation = (item.data(Qt.UserRole + 1) + degrees) % 360
        item.setData(Qt.UserRole + 1, rotation)
        original = self._pixmaps[item.data(Qt.UserRole)]
        item.setIcon(QIcon(original.transformed(QTransform().rotate(rotation), Qt.SmoothTransformation)))
        item.setText(f"Trang {item.data(Qt.UserRole) + 1}" + (f" · xoay {rotation}°" if rotation else ""))

    def _remove(self) -> None:
        if self.pages.count() <= 1:
            self.status.setText("Cần giữ lại ít nhất một trang.")
            return
        self.pages.takeItem(self.pages.currentRow())
        self.status.setText(f"Còn {self.pages.count()} trang. Trang bị bỏ chỉ bị loại khỏi tệp mới.")

    def _save(self) -> None:
        specification = tuple(
            (self.pages.item(row).data(Qt.UserRole), self.pages.item(row).data(Qt.UserRole + 1))
            for row in range(self.pages.count())
        )
        self.save.setEnabled(False)
        self.left.setEnabled(False)
        self.right.setEnabled(False)
        self.remove.setEnabled(False)
        self.pages.setEnabled(False)
        self.status.setText("Đang lưu PDF mới…")
        self._save_worker = SaveWorker(self.source, specification, self.output_dir)
        self._save_worker.succeeded.connect(self._saved)
        self._save_worker.failed.connect(self._save_failed)
        self._save_worker.start()

    def _saved(self, path: Path) -> None:
        if self._save_worker:
            self._save_worker.wait()
        self.saved_path = path
        self.accept()

    def _save_failed(self, message: str) -> None:
        if self._save_worker:
            self._save_worker.wait()
        self.status.setText("Không lưu được PDF: " + message)
        self._save_worker = None
        self.pages.setEnabled(True)
        self.save.setEnabled(True)
        self._selection_changed(self.pages.currentItem(), None)

    def closeEvent(self, event) -> None:
        if self._save_worker and self._save_worker.isRunning():
            self.status.setText("Đang lưu PDF, vui lòng đợi hoàn tất…")
            event.ignore()
            return
        if self._thumbnail_worker.isRunning():
            self._thumbnail_worker.requestInterruption()
            self._thumbnail_worker.wait()
        super().closeEvent(event)

    def done(self, result: int) -> None:
        if self._save_worker and self._save_worker.isRunning():
            self.status.setText("Đang lưu PDF, vui lòng đợi hoàn tất…")
            return
        if self._thumbnail_worker.isRunning():
            self._thumbnail_worker.requestInterruption()
            self._thumbnail_worker.wait()
        super().done(result)

