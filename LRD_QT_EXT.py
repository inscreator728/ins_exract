import sys
import os
import re
import time
import pandas as pd
import easyocr

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTextEdit, QMessageBox, QLabel, QProgressBar,
    QGroupBox, QRadioButton, QInputDialog, QPlainTextEdit, QMenuBar, QMenu
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QTimer
from PyQt6.QtGui import QIcon, QPixmap, QImage

# ------------------------- Globals -------------------------
text_length_limit = None
use_full_path = True


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name)


# ------------------------- Worker -------------------------
class OCRWorker(QThread):
    progress_step = pyqtSignal(int, int)   # processed, total
    file_done = pyqtSignal(str, str)
    file_preview = pyqtSignal(str)
    all_done = pyqtSignal(list, str)
    error_msg = pyqtSignal(str)
    log_msg = pyqtSignal(str)

    def __init__(self, folder_path: str, scan_mode: str, specific_files=None):
        super().__init__()
        self.folder_path = folder_path
        self.scan_mode = scan_mode
        self.specific_files = specific_files
        self.start_time = None

    def _get_reader(self):
        try:
            return easyocr.Reader(["en"], gpu=False)  # Force CPU only
        except Exception as e:
            self.error_msg.emit(f"Failed to load OCR model: {e}")
            return None

    def _extract_text(self, reader, image_path: str) -> str:
        try:
            results = reader.readtext(image_path)
            text = " ".join([r[1] for r in results])
            text = " ".join(text.split())
            global text_length_limit
            if text_length_limit is not None:
                text = text[:text_length_limit]
            return text.strip() if text.strip() else "[No text found]"
        except Exception as e:
            return f"Error: {e}"

    def run(self):
        try:
            if self.specific_files:
                files = self.specific_files
            else:
                files = [f for f in os.listdir(self.folder_path)
                         if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff"))]
        except Exception as e:
            self.error_msg.emit(f"Invalid folder: {e}")
            return

        total_files = len(files)
        if total_files == 0:
            self.error_msg.emit("No valid image files found in the selected folder.")
            return

        output_folder = os.path.join(self.folder_path, "Extracted_Texts")
        os.makedirs(output_folder, exist_ok=True)

        reader = self._get_reader()
        if reader is None:
            return

        records = []
        self.start_time = time.time()

        for idx, file_name in enumerate(files):
            if self.isInterruptionRequested():
                break

            file_path = os.path.join(self.folder_path, file_name)
            self.file_preview.emit(file_path)
            self.log_msg.emit(f"Processing: {file_name}")

            extracted_text = self._extract_text(reader, file_path)

            # --- Create folder for extracted text ---
            if extracted_text.strip() and not extracted_text.startswith("Error"):
                folder_name = sanitize_filename(extracted_text[:50].replace(" ", "_")) or "Extracted"
                folder_path_final = os.path.join(output_folder, folder_name)
                os.makedirs(folder_path_final, exist_ok=True)

                text_file_path = os.path.join(folder_path_final, "extracted_text.txt")
                with open(text_file_path, "w", encoding="utf-8") as text_file:
                    text_file.write(extracted_text)

                saved_path = folder_path_final if use_full_path else os.path.relpath(folder_path_final, self.folder_path)
            else:
                saved_path = "[No folder created]"

            # --- Append unified record ---
            records.append({
                "File Name": file_name,
                "Extracted Text": extracted_text,
                "Saved Path": saved_path
            })

            self.file_done.emit(file_path, extracted_text)
            self.progress_step.emit(idx + 1, total_files)

        if records:
            try:
                excel_file_path = os.path.join(output_folder, "extracted_texts.xlsx")
                pd.DataFrame(records).to_excel(excel_file_path, index=False)
                self.log_msg.emit(f"Saved results to {excel_file_path}")
            except Exception as e:
                self.error_msg.emit(f"Failed writing Excel: {e}")

        self.all_done.emit(records, output_folder)


# ------------------------- Main App -------------------------
class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LRD Text Extractor - PyQt6 v2.3 (Unified Save Mode)")
        self.resize(1100, 750)
        try:
            self.setWindowIcon(QIcon("Ticon.ico"))
        except Exception:
            pass

        self.scan_mode = "normal"
        self.worker: OCRWorker | None = None
        self.start_time = None
        self.total_files = 0
        self.processed_files = 0

        # ---------------- Menu Bar ----------------
        menubar = QMenuBar(self)
        help_menu = QMenu("Help", self)
        help_menu.addAction("About", self.show_about)
        menubar.addMenu(help_menu)
        self.setMenuBar(menubar)

        # --- Scan Mode Group ---
        scan_group = QGroupBox("Scan Mode")
        scan_layout = QHBoxLayout()
        self.rb_scan_normal = QRadioButton("Normal")
        self.rb_scan_super = QRadioButton("Super")
        self.rb_scan_intense = QRadioButton("Intense")
        self.rb_scan_normal.setChecked(True)
        for rb in [self.rb_scan_normal, self.rb_scan_super, self.rb_scan_intense]:
            scan_layout.addWidget(rb)
        scan_group.setLayout(scan_layout)

        # --- Buttons row ---
        self.btn_extract_single = QPushButton("Extract Single Image")
        self.btn_extract_bulk = QPushButton("Extract Bulk Images")
        self.btn_cancel = QPushButton("Cancel Process")
        self.btn_set_limit = QPushButton("Set Text Limit")

        top_buttons = QHBoxLayout()
        for b in [self.btn_extract_single, self.btn_extract_bulk,
                  self.btn_cancel, self.btn_set_limit]:
            top_buttons.addWidget(b)

        # --- Progress ---
        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progress: 0/0 (0%) | Time Left: --")

        # --- Previews ---
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)

        self.image_label = QLabel("Image preview")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setMinimumSize(400, 300)
        self.image_label.setStyleSheet("border: 1px solid #ccc;")

        preview_row = QHBoxLayout()
        preview_row.addWidget(self.text_preview, stretch=1)
        preview_row.addWidget(self.image_label, stretch=1)

        # --- Log Panel ---
        self.log_panel = QPlainTextEdit()
        self.log_panel.setReadOnly(True)

        # --- Layout root ---
        root_layout = QVBoxLayout()
        root_layout.addWidget(scan_group)
        root_layout.addLayout(top_buttons)
        root_layout.addWidget(self.progress_bar)
        root_layout.addWidget(self.progress_label)
        root_layout.addLayout(preview_row)
        root_layout.addWidget(QLabel("Logs:"))
        root_layout.addWidget(self.log_panel)

        container = QWidget()
        container.setLayout(root_layout)
        self.setCentralWidget(container)

        # --- Connect ---
        self.rb_scan_normal.toggled.connect(lambda c: c and self._set_scan_mode("normal"))
        self.rb_scan_super.toggled.connect(lambda c: c and self._set_scan_mode("super"))
        self.rb_scan_intense.toggled.connect(lambda c: c and self._set_scan_mode("intense"))

        self.btn_extract_single.clicked.connect(self.extract_single)
        self.btn_extract_bulk.clicked.connect(self.extract_bulk)
        self.btn_cancel.clicked.connect(self.cancel_process)
        self.btn_set_limit.clicked.connect(self.set_text_limit)

        # --- Timer for live clock ---
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_clock)

    # ---------- Config setters ----------
    def _set_scan_mode(self, mode: str):
        self.scan_mode = mode
        self.log_panel.appendPlainText(f"Scan mode set to: {mode}")

    # ---------- Actions ----------
    def set_text_limit(self):
        global text_length_limit
        current = text_length_limit if text_length_limit is not None else 0
        val, ok = QInputDialog.getInt(
            self, "Set Text Limit", "Enter maximum characters (0 = No Limit):",
            value=current, min=0
        )
        if ok:
            text_length_limit = None if val == 0 else val
            self.log_panel.appendPlainText(
                f"Text limit set to: {'No Limit' if text_length_limit is None else text_length_limit}"
            )

    def extract_single(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.tiff)"
        )
        if not file_path:
            return
        folder_path = os.path.dirname(file_path)
        self._run_worker(folder_path, [os.path.basename(file_path)])

    def extract_bulk(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if not folder_path:
            return
        self._run_worker(folder_path)

    def _run_worker(self, folder_path, specific_files=None):
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "Busy", "Process already running.")
            return
        self.text_preview.clear()
        self.image_label.clear()
        self.log_panel.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("Progress: 0/0 (0%) | Time Left: --")

        self.worker = OCRWorker(folder_path, self.scan_mode, specific_files)
        self._connect_worker_signals()
        self.worker.start()

        self.start_time = time.time()
        self.processed_files = 0
        self.total_files = 0
        self.timer.start(1000)  # tick every second

    def _connect_worker_signals(self):
        self.worker.progress_step.connect(self._on_progress)
        self.worker.file_done.connect(self._on_file_done)
        self.worker.file_preview.connect(self._on_file_preview)
        self.worker.all_done.connect(self._on_all_done)
        self.worker.error_msg.connect(self._on_error)
        self.worker.log_msg.connect(self._on_log)

    # ---------- Signal handlers ----------
    def _on_progress(self, processed: int, total: int):
        self.processed_files = processed
        self.total_files = total
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(processed)
        self.update_clock()

    def update_clock(self):
        if not self.start_time or self.total_files == 0:
            return
        elapsed = time.time() - self.start_time
        avg_time = elapsed / self.processed_files if self.processed_files > 0 else 0
        remaining = (self.total_files - self.processed_files) * avg_time
        mins, secs = divmod(int(remaining), 60)
        pct = (self.processed_files / self.total_files) * 100 if self.total_files else 0
        self.progress_label.setText(
            f"Progress: {self.processed_files}/{self.total_files} ({pct:.2f}%) | Time Left: {mins}m {secs}s"
        )

    def _on_file_done(self, file_path: str, text: str):
        self.text_preview.clear()
        self.text_preview.append(text if text else "[No text found]")

    def _on_file_preview(self, file_path: str):
        try:
            img = QImage(file_path)
            if not img.isNull():
                pix = QPixmap.fromImage(img).scaled(
                    400, 300, Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.image_label.setPixmap(pix)
            else:
                self.image_label.clear()
        except Exception:
            self.image_label.clear()

    def _on_all_done(self, records: list, output_folder: str):
        self.timer.stop()
        self.log_panel.appendPlainText("Extraction finished.")
        QMessageBox.information(
            self, "Success", f"Results saved in:\n{output_folder}"
        )
        try:
            if sys.platform.startswith("win"):
                os.startfile(output_folder)
            elif sys.platform == "darwin":
                os.system(f'open "{output_folder}"')
            else:
                os.system(f'xdg-open "{output_folder}"')
        except Exception:
            pass

    def _on_error(self, msg: str):
        self.timer.stop()
        self.log_panel.appendPlainText(f"ERROR: {msg}")
        QMessageBox.critical(self, "Error", msg)

    def _on_log(self, msg: str):
        self.log_panel.appendPlainText(msg)

    def cancel_process(self):
        if self.worker and self.worker.isRunning():
            self.worker.requestInterruption()
            self.timer.stop()
            self.log_panel.appendPlainText("Process cancellation requested.")

    def show_about(self):
        QMessageBox.information(
            self, "About LRD Text Extractor",
            "LRD Text Extractor v2.3 (Unified Save Mode)\n\n"
            "✔ Extract text from images/folders\n"
            "✔ Save results into Excel automatically\n"
            "✔ Each extracted text also saved in named folder\n"
            "✔ CPU-only (runs even on i3 / 2GB RAM / no GPU)\n\n"
            "Developer: LRD_SOUL\n"
            "Telegram: @LRD_SOUL\n"
        )


# ------------------------- Run -------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OCRApp()
    window.show()
    sys.exit(app.exec())

