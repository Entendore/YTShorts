#!/usr/bin/env python3
"""
PPTX to MP4 Converter 
Converts PowerPoint (.pptx) presentations to MP4 video with optional background music.

Features:
  - Project folder with named projects (each a self-contained subfolder)
  - Project list panel with cards (name, date, thumbnail, slide count)
  - Project CRUD: create, rename, duplicate, delete, open in file manager
  - Slide preview gallery with cached thumbnails per project
  - Drag-and-drop PPTX import
  - Background music with volume control

Project folder structure:
    ~/PPTX2MP4_Projects/
      My Presentation/
        project.json       # settings
        slides/
          slide_0001.png   # cached thumbnails
          slide_0002.png
          ...

Dependencies:
    pip install PySide6 python-pptx Pillow
    System:  libreoffice, ffmpeg
    Optional: poppler-utils (for pdf2image, better thumbnail quality)
"""

import json
import os
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QPushButton, QLineEdit, QFileDialog,
    QSpinBox, QDoubleSpinBox, QComboBox, QProgressBar, QMessageBox,
    QScrollArea, QSplitter, QListWidget, QListWidgetItem, QMenu,
    QSizePolicy, QInputDialog, QAbstractItemView, QHeaderView,
    QTableWidget, QTableWidgetItem, QToolButton, QFrame,
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QSize, QDateTime,
)
from PySide6.QtGui import (
    QFont, QPalette, QColor, QPixmap, QIcon, QAction,
    QDragEnterEvent, QDropEvent, QPainter, QPainterPath,
)

# ═══════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════

DEFAULT_PROJECTS_DIR = os.path.expanduser("~/PPTX2MP4_Projects")
SETTINGS_FILE = os.path.expanduser("~/.pptx2mp4_settings.json")
PROJECT_JSON = "project.json"
SLIDES_DIR = "slides"

RESOLUTION_PRESETS = {
    "1920 x 1080  (Full HD)": (1920, 1080),
    "1280 x 720   (HD)":      (1280, 720),
    "3840 x 2160  (4K)":      (3840, 2160),
    "854 x 480    (SD)":      (854, 480),
    "Custom...":              None,
}


# ═══════════════════════════════════════════════════════════════════════
#  Thumbnail generation worker
# ═══════════════════════════════════════════════════════════════════════

class ThumbnailWorker(QThread):
    """Converts a PPTX file to slide thumbnail PNGs."""

    progress = Signal(int, str)
    thumbnail_ready = Signal(int, str)
    all_done = Signal(int)
    error = Signal(str)

    def __init__(self, pptx_path: str, output_dir: str, thumbnail_dpi: int = 150):
        super().__init__()
        self.pptx_path = pptx_path
        self.output_dir = output_dir
        self.thumbnail_dpi = thumbnail_dpi
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        tmpdir = None
        try:
            os.makedirs(self.output_dir, exist_ok=True)

            # Clear old thumbnails
            for old in Path(self.output_dir).glob("*.png"):
                old.unlink()

            import pptx
            prs = pptx.Presentation(self.pptx_path)
            n = len(prs.slides)
            self.progress.emit(5, f"Presentation has {n} slide(s)")
            if self._cancel:
                return

            # PPTX -> PDF
            tmpdir = tempfile.mkdtemp(prefix="pptx_thumb_")
            pdf_dir = os.path.join(tmpdir, "pdf")
            os.makedirs(pdf_dir, exist_ok=True)
            cmd = [
                "libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", pdf_dir, self.pptx_path,
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if result.returncode != 0:
                raise RuntimeError(
                    f"LibreOffice failed (exit {result.returncode}):\n"
                    f"{result.stderr or result.stdout}"
                )

            pdf_files = list(Path(pdf_dir).glob("*.pdf"))
            if not pdf_files:
                raise RuntimeError("LibreOffice did not produce a PDF")
            pdf_path = str(pdf_files[0])
            if self._cancel:
                return

            # PDF -> PNG
            self.progress.emit(30, "Generating thumbnails...")
            count = 0
            try:
                from pdf2image import convert_from_path
                images = convert_from_path(pdf_path, dpi=self.thumbnail_dpi)
                for i, img in enumerate(images):
                    if self._cancel:
                        return
                    out = os.path.join(self.output_dir, f"slide_{i + 1:04d}.png")
                    img.save(out, "PNG")
                    self.thumbnail_ready.emit(i, out)
                    count += 1
                    self.progress.emit(30 + int((i + 1) / len(images) * 60),
                                       f"Thumbnail {i + 1}/{len(images)}")
            except ImportError:
                cmd = [
                    "libreoffice", "--headless", "--convert-to", "png",
                    "--outdir", self.output_dir, pdf_path,
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                if result.returncode != 0:
                    raise RuntimeError(f"PNG conversion failed:\n{result.stderr}")
                pngs = sorted(Path(self.output_dir).glob("*.png"))
                for i, pf in enumerate(pngs):
                    if self._cancel:
                        return
                    target = os.path.join(self.output_dir, f"slide_{i + 1:04d}.png")
                    if pf.resolve() != Path(target).resolve():
                        shutil.move(str(pf), target)
                    self.thumbnail_ready.emit(i, target)
                    count += 1
                self.progress.emit(90, "Thumbnails ready")

            self.progress.emit(100, "Done")
            self.all_done.emit(count)

        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)


# ═══════════════════════════════════════════════════════════════════════
#  Conversion worker
# ═══════════════════════════════════════════════════════════════════════

class ConversionWorker(QThread):
    progress = Signal(int)
    status_message = Signal(str)
    finished_ok = Signal(str)
    finished_error = Signal(str)

    def __init__(self, pptx_path, output_path, music_path, seconds_per_slide,
                 fps, resolution, quality):
        super().__init__()
        self.pptx_path = pptx_path
        self.output_path = output_path
        self.music_path = music_path
        self.seconds_per_slide = seconds_per_slide
        self.fps = fps
        self.resolution = resolution
        self.quality = quality
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def _run_cmd(self, cmd, label=""):
        self.status_message.emit(f"Running: {label}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        if result.returncode != 0:
            raise RuntimeError(
                f"{label} failed (exit {result.returncode}):\n"
                f"{result.stderr or result.stdout}"
            )

    def run(self):
        tmpdir = None
        try:
            if self._cancel:
                return
            tmpdir = tempfile.mkdtemp(prefix="pptx2mp4_")
            self.status_message.emit("Preparing...")

            import pptx
            prs = pptx.Presentation(self.pptx_path)
            slide_count = len(prs.slides)
            self.status_message.emit(f"{slide_count} slide(s)")

            self.progress.emit(5)
            self.status_message.emit("Converting slides (LibreOffice)...")
            pdf_dir = os.path.join(tmpdir, "pdf")
            os.makedirs(pdf_dir, exist_ok=True)
            self._run_cmd([
                "libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", pdf_dir, self.pptx_path,
            ], "LibreOffice PPTX -> PDF")

            pdf_files = list(Path(pdf_dir).glob("*.pdf"))
            if not pdf_files:
                raise RuntimeError("LibreOffice did not produce a PDF")
            pdf_path = str(pdf_files[0])

            output_dir = os.path.join(tmpdir, "slides")
            os.makedirs(output_dir, exist_ok=True)

            try:
                from pdf2image import convert_from_path
                self.status_message.emit("PDF -> PNG (pdf2image)...")
                images = convert_from_path(pdf_path, dpi=200)
                for i, img in enumerate(images, start=1):
                    img.save(os.path.join(output_dir, f"slide_{i:04d}.png"), "PNG")
                    self.progress.emit(5 + int((i / len(images)) * 40))
                png_count = len(images)
            except ImportError:
                self.status_message.emit("PDF -> PNG (LibreOffice)...")
                self._run_cmd([
                    "libreoffice", "--headless", "--convert-to", "png",
                    "--outdir", output_dir, pdf_path,
                ], "LibreOffice PDF -> PNG")
                pngs = sorted(Path(output_dir).glob("*.png"))
                for i, pf in enumerate(pngs, start=1):
                    target = os.path.join(output_dir, f"slide_{i:04d}.png")
                    if pf.resolve() != Path(target).resolve():
                        shutil.move(str(pf), target)
                png_count = len(pngs)
                self.progress.emit(45)

            if self._cancel:
                return

            slide_images = sorted(Path(output_dir).glob("*.png"))
            if not slide_images:
                raise RuntimeError("No slide images generated.")
            self.status_message.emit(f"Found {len(slide_images)} slides")

            concat_path = os.path.join(tmpdir, "concat.txt")
            with open(concat_path, "w") as f:
                for img_path in slide_images:
                    f.write(f"file '{img_path}'\n")
                    f.write(f"duration {self.seconds_per_slide}\n")
                f.write(f"file '{slide_images[-1]}'\n")

            w, h = self.resolution
            cmd = ["ffmpeg", "-y"]
            cmd += ["-f", "concat", "-safe", "0", "-i", concat_path]
            has_music = self.music_path and os.path.isfile(self.music_path)
            if has_music:
                cmd += ["-i", self.music_path]
            cmd += [
                "-c:v", "libx264", "-pix_fmt", "yuv420p",
                "-vf", (
                    f"scale={w}:{h}:force_original_aspect_ratio=decrease,"
                    f"pad={w}:{h}:(ow-iw)/2:(oh-ih)/2:color=black,fps={self.fps}"
                ),
                "-q:v", str(self.quality),
            ]
            if has_music:
                cmd += ["-c:a", "aac", "-b:a", "192k", "-shortest"]
            else:
                cmd += ["-an"]
            cmd.append(self.output_path)

            self._run_ffmpeg(cmd, slide_count)
            if self._cancel:
                return
            self.progress.emit(100)
            self.status_message.emit("Done!")
            self.finished_ok.emit(self.output_path)

        except Exception as exc:
            self.finished_error.emit(str(exc))
        finally:
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

    def _run_ffmpeg(self, cmd, num_slides):
        self.status_message.emit("Encoding MP4 (ffmpeg)...")
        process = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1,
        )
        total_dur = num_slides * self.seconds_per_slide
        while True:
            line = process.stdout.readline()
            if not line and process.poll() is not None:
                break
            if self._cancel:
                process.terminate()
                return
            if "time=" in line:
                try:
                    ts = line.split("time=")[1].split(" ")[0]
                    p = ts.split(":")
                    t = int(p[0]) * 3600 + int(p[1]) * 60 + float(p[2])
                    if total_dur > 0:
                        self.progress.emit(min(int((t / total_dur) * 50) + 50, 99))
                except (ValueError, IndexError):
                    pass
        if process.returncode != 0:
            raise RuntimeError(
                f"ffmpeg failed (exit {process.returncode}):\n{process.stdout.read()}"
            )


# ═══════════════════════════════════════════════════════════════════════
#  Project card widget  (for the project list)
# ═══════════════════════════════════════════════════════════════════════

class ProjectCard(QFrame):
    """A single project card shown in the project list."""

    def __init__(self, name: str, project_dir: str, parent=None):
        super().__init__(parent)
        self.project_name = name
        self.project_dir = project_dir
        self.setFixedHeight(80)
        self._build()

    def _build(self):
        self.setStyleSheet("""
            ProjectCard {
                background-color: #313244;
                border: 1px solid #45475a;
                border-radius: 8px;
                margin: 2px;
            }
            ProjectCard:hover {
                border-color: #89b4fa;
                background-color: #373848;
            }
        """)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 8)

        # Thumbnail (first slide if available)
        self.thumb_label = QLabel()
        self.thumb_label.setFixedSize(56, 56)
        self.thumb_label.setAlignment(Qt.AlignCenter)
        self.thumb_label.setStyleSheet(
            "background-color: #181825; border-radius: 4px;"
        )
        # Try to load first-slide thumbnail
        first_thumb = os.path.join(self.project_dir, SLIDES_DIR, "slide_0001.png")
        if os.path.isfile(first_thumb):
            pm = QPixmap(first_thumb)
            if not pm.isNull():
                self.thumb_label.setPixmap(
                    pm.scaled(56, 56, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                )
        lay.addWidget(self.thumb_label)

        # Info column
        info = QVBoxLayout()
        info.setSpacing(2)

        self.lbl_name = QLabel(self.project_name)
        self.lbl_name.setStyleSheet(
            "color: #cdd6f4; font-size: 13px; font-weight: bold; border: none;"
        )
        info.addWidget(self.lbl_name)

        # Load project.json for metadata
        pj = os.path.join(self.project_dir, PROJECT_JSON)
        slide_count = 0
        date_str = ""
        pptx_name = ""
        if os.path.isfile(pj):
            try:
                with open(pj, "r", encoding="utf-8") as f:
                    data = json.load(f)
                pptx_name = os.path.basename(data.get("pptx_path", ""))
                slide_count = data.get("slide_count", 0)
                mtime = os.path.getmtime(pj)
                date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            except Exception:
                mtime = os.path.getmtime(self.project_dir)
                date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
        else:
            mtime = os.path.getmtime(self.project_dir)
            date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")

        self.lbl_detail = QLabel()
        parts = []
        if pptx_name:
            parts.append(pptx_name)
        if slide_count:
            parts.append(f"{slide_count} slides")
        parts.append(date_str)
        self.lbl_detail.setText("  |  ".join(parts))
        self.lbl_detail.setStyleSheet(
            "color: #a6adc8; font-size: 11px; border: none;"
        )
        info.addWidget(self.lbl_detail)
        lay.addLayout(info, 1)

        # Open folder button
        self.btn_folder = QToolButton()
        self.btn_folder.setText("...")
        self.btn_folder.setFixedSize(28, 28)
        self.btn_folder.setToolTip("Open project folder")
        self.btn_folder.setStyleSheet(
            "QToolButton { background-color: #45475a; color: #cdd6f4; "
            "border: 1px solid #585b70; border-radius: 4px; font-size: 14px; }"
            "QToolButton:hover { background-color: #585b70; }"
        )
        lay.addWidget(self.btn_folder)


# ═══════════════════════════════════════════════════════════════════════
#  Slide thumbnail widget  (for the slide preview list)
# ═══════════════════════════════════════════════════════════════════════

class SlideThumbWidget(QWidget):
    """Single slide thumbnail card inside the preview list."""

    def __init__(self, index: int, image_path: str, parent=None):
        super().__init__(parent)
        self.index = index
        self.image_path = image_path
        self.setFixedSize(180, 140)
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(4, 4, 4, 4)
        lay.setSpacing(4)
        self.lbl_img = QLabel()
        self.lbl_img.setAlignment(Qt.AlignCenter)
        self.lbl_img.setFixedSize(170, 110)
        self.lbl_img.setStyleSheet(
            "background-color: #181825; border: 1px solid #45475a; border-radius: 4px;"
        )
        pm = QPixmap(self.image_path)
        if not pm.isNull():
            self.lbl_img.setPixmap(
                pm.scaled(170, 110, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            )
        lay.addWidget(self.lbl_img, 1)
        self.lbl_name = QLabel(f"Slide {self.index + 1}")
        self.lbl_name.setAlignment(Qt.AlignCenter)
        self.lbl_name.setStyleSheet("color: #a6adc8; font-size: 11px; border: none;")
        lay.addWidget(self.lbl_name)


# ═══════════════════════════════════════════════════════════════════════
#  Stylesheet
# ═══════════════════════════════════════════════════════════════════════

STYLESHEET = """
QMainWindow { background-color: #1e1e2e; }
QMenuBar {
    background-color: #181825; color: #cdd6f4;
    border-bottom: 1px solid #313244;
}
QMenuBar::item { padding: 6px 14px; }
QMenuBar::item:selected { background-color: #45475a; border-radius: 4px; }
QMenu {
    background-color: #313244; color: #cdd6f4;
    border: 1px solid #45475a; padding: 4px;
}
QMenu::item { padding: 6px 32px 6px 14px; border-radius: 4px; }
QMenu::item:selected { background-color: #89b4fa; color: #1e1e2e; }
QMenu::separator { height: 1px; background: #45475a; margin: 4px 8px; }
QToolBar {
    background-color: #181825; border-bottom: 1px solid #313244;
    spacing: 6px; padding: 4px;
}
QGroupBox {
    color: #cdd6f4; font-size: 13px; font-weight: bold;
    border: 1px solid #45475a; border-radius: 8px;
    margin-top: 10px; padding: 14px 10px 10px 10px;
    background-color: #313244;
}
QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 0 6px; }
QLabel { color: #cdd6f4; font-size: 13px; }
QLineEdit {
    background-color: #45475a; color: #cdd6f4;
    border: 1px solid #585b70; border-radius: 6px;
    padding: 5px 10px; font-size: 13px;
}
QLineEdit:read-only { background-color: #383850; color: #a6adc8; }
QPushButton {
    background-color: #89b4fa; color: #1e1e2e;
    border: none; border-radius: 6px; padding: 7px 14px;
    font-size: 13px; font-weight: bold;
}
QPushButton:hover { background-color: #b4d0fb; }
QPushButton:pressed { background-color: #74c7ec; }
QPushButton:disabled { background-color: #585b70; color: #6c7086; }
QPushButton#convertBtn {
    background-color: #a6e3a1; font-size: 15px; padding: 10px 20px;
}
QPushButton#convertBtn:hover { background-color: #b8f0b4; }
QPushButton#convertBtn:disabled { background-color: #585b70; color: #6c7086; }
QPushButton#cancelBtn { background-color: #f38ba8; color: #1e1e2e; }
QPushButton#cancelBtn:hover { background-color: #f5a3ba; }
QPushButton#browseMusicBtn { background-color: #cba6f7; color: #1e1e2e; }
QPushButton#browseMusicBtn:hover { background-color: #d4bafb; }
QPushButton#newProjBtn { background-color: #a6e3a1; color: #1e1e2e; }
QPushButton#newProjBtn:hover { background-color: #b8f0b4; }
QPushButton#importBtn { background-color: #89b4fa; color: #1e1e2e; }
QPushButton#importBtn:hover { background-color: #b4d0fb; }
QPushButton#deleteProjBtn { background-color: #f38ba8; color: #1e1e2e; }
QPushButton#deleteProjBtn:hover { background-color: #f5a3ba; }
QPushButton#renameProjBtn { background-color: #f9e2af; color: #1e1e2e; }
QPushButton#renameProjBtn:hover { background-color: #fbe8c4; }
QSpinBox, QDoubleSpinBox, QComboBox {
    background-color: #45475a; color: #cdd6f4;
    border: 1px solid #585b70; border-radius: 6px;
    padding: 5px 10px; font-size: 13px;
}
QComboBox::drop-down { border: none; width: 28px; }
QComboBox QAbstractItemView {
    background-color: #45475a; color: #cdd6f4;
    selection-background-color: #89b4fa;
}
QProgressBar {
    background-color: #45475a; border: none; border-radius: 6px;
    text-align: center; color: #cdd6f4; font-size: 12px; min-height: 20px;
}
QProgressBar::chunk { background-color: #a6e3a1; border-radius: 6px; }
QStatusBar { color: #a6adc8; font-size: 12px; background-color: #181825; }
QScrollArea { background-color: #1e1e2e; border: none; }
QListWidget {
    background-color: #181825; color: #cdd6f4; border: none;
    border-radius: 8px; outline: none;
}
QListWidget#projectList {
    border: 1px solid #45475a; border-radius: 8px;
    background-color: #1e1e2e;
}
QListWidget#projectList::item { padding: 3px; margin: 2px 0px; border-radius: 0px; }
QListWidget#projectList::item:selected {
    background-color: transparent;
}
QListWidget::item { padding: 4px; margin: 2px; border-radius: 6px; }
QListWidget::item:selected { background-color: #45475a; }
QListWidget::item:hover { background-color: #313244; }
QSplitter::handle { background-color: #45475a; }
QSplitter::handle:vertical { height: 3px; }
QSplitter::handle:horizontal { width: 3px; }
QCheckBox { color: #cdd6f4; font-size: 13px; }
QCheckBox::indicator {
    width: 18px; height: 18px; border-radius: 4px;
    border: 1px solid #585b70; background-color: #45475a;
}
QCheckBox::indicator:checked { background-color: #89b4fa; border-color: #89b4fa; }
QToolButton {
    background-color: #45475a; color: #cdd6f4;
    border: 1px solid #585b70; border-radius: 4px;
    padding: 4px 8px; font-size: 13px;
}
QToolButton:hover { background-color: #585b70; }
"""


# ═══════════════════════════════════════════════════════════════════════
#  Main Window
# ═══════════════════════════════════════════════════════════════════════

class PPTXToMP4Converter(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPTX to MP4 Converter")
        self.setMinimumSize(1000, 680)
        self.resize(1160, 780)

        self.worker: ConversionWorker | None = None
        self.thumb_worker: ThumbnailWorker | None = None

        # Project state
        self.projects_dir = DEFAULT_PROJECTS_DIR
        self._current_project_dir: str | None = None
        self._current_project_name: str | None = None
        self._current_pptx: str = ""
        self._slide_count: int = 0

        os.makedirs(self.projects_dir, exist_ok=True)

        self._build_menubar()
        self._build_ui()
        self._refresh_project_list()
        self._restore_window_settings()
        self.setAcceptDrops(True)

    # ── Menu bar ──────────────────────────────────────────────────

    def _build_menubar(self):
        mb = self.menuBar()

        file_menu = mb.addMenu("&File")

        act_new = QAction("&New Project", self)
        act_new.setShortcut("Ctrl+N")
        act_new.triggered.connect(self._new_project)
        file_menu.addAction(act_new)

        act_import = QAction("&Import PPTX...", self)
        act_import.setShortcut("Ctrl+I")
        act_import.triggered.connect(self._browse_pptx)
        file_menu.addAction(act_import)

        file_menu.addSeparator()

        act_rename = QAction("&Rename Project", self)
        act_rename.setShortcut("F2")
        act_rename.triggered.connect(self._rename_project)
        file_menu.addAction(act_rename)

        act_dup = QAction("&Duplicate Project", self)
        act_dup.triggered.connect(self._duplicate_project)
        file_menu.addAction(act_dup)

        act_del = QAction("&Delete Project", self)
        act_del.triggered.connect(self._delete_project)
        file_menu.addAction(act_del)

        file_menu.addSeparator()

        act_folder = QAction("Open Project &Folder", self)
        act_folder.triggered.connect(self._open_project_folder)
        file_menu.addAction(act_folder)

        act_projdir = QAction("Set Projects &Directory...", self)
        act_projdir.triggered.connect(self._change_projects_dir)
        file_menu.addAction(act_projdir)

        file_menu.addSeparator()

        act_exit = QAction("E&xit", self)
        act_exit.setShortcut("Ctrl+Q")
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

        help_menu = mb.addMenu("&Help")
        act_about = QAction("&About", self)
        act_about.triggered.connect(self._show_about)
        help_menu.addAction(act_about)

    # ── Central UI ────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 4, 8, 8)
        root.setSpacing(0)

        # Three-panel horizontal splitter
        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, 1)

        # ── Panel 1: Project List (left) ──
        panel_projects = QWidget()
        lay_proj = QVBoxLayout(panel_projects)
        lay_proj.setContentsMargins(0, 0, 4, 0)
        lay_proj.setSpacing(6)

        proj_header = QHBoxLayout()
        proj_title = QLabel("Projects")
        proj_title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        proj_title.setStyleSheet("color: #cdd6f4;")
        proj_header.addWidget(proj_title)
        self.lbl_proj_dir = QLabel()
        self.lbl_proj_dir.setStyleSheet("color: #6c7086; font-size: 10px;")
        self.lbl_proj_dir.setText(f"  {self.projects_dir}")
        proj_header.addStretch()
        lay_proj.addLayout(proj_header)

        # Action buttons row
        proj_btns = QHBoxLayout()
        self.btn_new_proj = QPushButton("+ New")
        self.btn_new_proj.setObjectName("newProjBtn")
        self.btn_new_proj.clicked.connect(self._new_project)
        proj_btns.addWidget(self.btn_new_proj)

        self.btn_rename_proj = QPushButton("Rename")
        self.btn_rename_proj.setObjectName("renameProjBtn")
        self.btn_rename_proj.clicked.connect(self._rename_project)
        self.btn_rename_proj.setEnabled(False)
        proj_btns.addWidget(self.btn_rename_proj)

        self.btn_delete_proj = QPushButton("Delete")
        self.btn_delete_proj.setObjectName("deleteProjBtn")
        self.btn_delete_proj.clicked.connect(self._delete_project)
        self.btn_delete_proj.setEnabled(False)
        proj_btns.addWidget(self.btn_delete_proj)
        lay_proj.addLayout(proj_btns)

        # Project list
        self.project_list = QListWidget()
        self.project_list.setObjectName("projectList")
        self.project_list.setSpacing(2)
        self.project_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.project_list.currentRowChanged.connect(self._on_project_selected)
        self.project_list.customContextMenuRequested.connect(self._on_project_context_menu)
        self.project_list.itemDoubleClicked.connect(lambda item: self._load_selected_project())
        lay_proj.addWidget(self.project_list, 1)

        self.lbl_proj_count = QLabel("0 projects")
        self.lbl_proj_count.setStyleSheet("color: #6c7086; font-size: 11px;")
        lay_proj.addWidget(self.lbl_proj_count)

        splitter.addWidget(panel_projects)

        # ── Panel 2: Slide Preview (middle) ──
        panel_slides = QWidget()
        lay_slides = QVBoxLayout(panel_slides)
        lay_slides.setContentsMargins(4, 0, 4, 0)
        lay_slides.setSpacing(6)

        slide_header = QHBoxLayout()
        slide_title = QLabel("Slide Preview")
        slide_title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        slide_title.setStyleSheet("color: #cdd6f4;")
        slide_header.addWidget(slide_title)
        slide_header.addStretch()

        self.btn_import = QPushButton("Import PPTX...")
        self.btn_import.setObjectName("importBtn")
        self.btn_import.clicked.connect(self._browse_pptx)
        slide_header.addWidget(self.btn_import)
        lay_slides.addLayout(slide_header)

        self.slide_list = QListWidget()
        self.slide_list.setViewMode(QListWidget.IconMode)
        self.slide_list.setIconSize(QSize(180, 130))
        self.slide_list.setResizeMode(QListWidget.Adjust)
        self.slide_list.setSpacing(8)
        self.slide_list.setWrapping(True)
        self.slide_list.setSelectionMode(QListWidget.SingleSelection)
        self.slide_list.setStyleSheet(
            "QListWidget { background-color: #181825; "
            "border: 2px dashed #45475a; border-radius: 10px; padding: 10px; }"
        )
        self._show_empty_slides()
        lay_slides.addWidget(self.slide_list, 1)

        self.lbl_slide_count = QLabel("")
        self.lbl_slide_count.setStyleSheet("color: #89b4fa; font-size: 11px;")
        lay_slides.addWidget(self.lbl_slide_count)

        splitter.addWidget(panel_slides)

        # ── Panel 3: Settings + Convert (right) ──
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        right_widget = QWidget()
        right_lay = QVBoxLayout(right_widget)
        right_lay.setContentsMargins(8, 4, 4, 4)
        right_lay.setSpacing(10)

        # Music
        grp_music = QGroupBox("Background Music (optional)")
        lay_music = QVBoxLayout(grp_music)
        row_music = QHBoxLayout()
        self.le_music = QLineEdit()
        self.le_music.setReadOnly(True)
        self.le_music.setPlaceholderText("Select audio (MP3, WAV, OGG)...")
        self.btn_music = QPushButton("Browse...")
        self.btn_music.setObjectName("browseMusicBtn")
        self.btn_music.clicked.connect(self._browse_music)
        self.btn_clear_music = QPushButton("X")
        self.btn_clear_music.setFixedWidth(32)
        self.btn_clear_music.setToolTip("Remove music")
        self.btn_clear_music.clicked.connect(self._clear_music)
        row_music.addWidget(self.le_music, 1)
        row_music.addWidget(self.btn_music)
        row_music.addWidget(self.btn_clear_music)
        lay_music.addLayout(row_music)
        row_vol = QHBoxLayout()
        row_vol.addWidget(QLabel("Volume:"))
        self.spin_volume = QSpinBox()
        self.spin_volume.setRange(0, 200)
        self.spin_volume.setValue(100)
        self.spin_volume.setSuffix("%")
        row_vol.addWidget(self.spin_volume)
        row_vol.addStretch()
        lay_music.addLayout(row_vol)
        right_lay.addWidget(grp_music)

        # Output
        grp_output = QGroupBox("Output Video")
        lay_output = QHBoxLayout(grp_output)
        self.le_output = QLineEdit()
        self.le_output.setReadOnly(True)
        self.le_output.setPlaceholderText("Output .mp4 path...")
        btn_out = QPushButton("Browse...")
        btn_out.clicked.connect(self._browse_output)
        lay_output.addWidget(self.le_output, 1)
        lay_output.addWidget(btn_out)
        right_lay.addWidget(grp_output)

        # Settings
        grp_settings = QGroupBox("Settings")
        lay_settings = QVBoxLayout(grp_settings)
        row_res = QHBoxLayout()
        row_res.addWidget(QLabel("Resolution:"))
        self.combo_res = QComboBox()
        self.combo_res.addItems(list(RESOLUTION_PRESETS.keys()))
        self.combo_res.currentIndexChanged.connect(self._on_resolution_changed)
        row_res.addWidget(self.combo_res, 1)
        lay_settings.addLayout(row_res)

        self.row_custom = QHBoxLayout()
        self.spin_w = QSpinBox()
        self.spin_w.setRange(320, 7680)
        self.spin_w.setValue(1920)
        self.spin_h = QSpinBox()
        self.spin_h.setRange(240, 4320)
        self.spin_h.setValue(1080)
        self.row_custom.addWidget(QLabel("W:"))
        self.row_custom.addWidget(self.spin_w)
        self.row_custom.addWidget(QLabel("H:"))
        self.row_custom.addWidget(self.spin_h)
        self.row_custom.addStretch()
        self.custom_widget = QWidget()
        self.custom_widget.setLayout(self.row_custom)
        self.custom_widget.hide()
        lay_settings.addWidget(self.custom_widget)

        row_dur = QHBoxLayout()
        row_dur.addWidget(QLabel("Seconds per slide:"))
        self.spin_dur = QDoubleSpinBox()
        self.spin_dur.setRange(0.5, 60.0)
        self.spin_dur.setValue(5.0)
        self.spin_dur.setSingleStep(0.5)
        self.spin_dur.setDecimals(1)
        row_dur.addWidget(self.spin_dur)
        row_dur.addStretch()
        lay_settings.addLayout(row_dur)

        row_fps = QHBoxLayout()
        row_fps.addWidget(QLabel("FPS:"))
        self.combo_fps = QComboBox()
        self.combo_fps.addItems(["24", "25", "30", "60"])
        self.combo_fps.setCurrentText("30")
        row_fps.addWidget(self.combo_fps)
        row_fps.addStretch()
        lay_settings.addLayout(row_fps)

        row_q = QHBoxLayout()
        row_q.addWidget(QLabel("Quality:"))
        self.combo_quality = QComboBox()
        self.combo_quality.addItems([
            "High  (larger file)", "Medium (balanced)", "Low   (smaller file)",
        ])
        self.combo_quality.setCurrentIndex(1)
        row_q.addWidget(self.combo_quality, 1)
        row_q.addStretch()
        lay_settings.addLayout(row_q)
        right_lay.addWidget(grp_settings)

        # Convert
        grp_conv = QGroupBox("Convert")
        lay_conv = QVBoxLayout(grp_conv)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        lay_conv.addWidget(self.progress_bar)
        self.lbl_status = QLabel("Select or create a project to begin.")
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet("color: #a6adc8; font-size: 12px;")
        lay_conv.addWidget(self.lbl_status)
        row_btns = QHBoxLayout()
        self.btn_convert = QPushButton("Convert to MP4")
        self.btn_convert.setObjectName("convertBtn")
        self.btn_convert.clicked.connect(self._start_conversion)
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setObjectName("cancelBtn")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_conversion)
        row_btns.addWidget(self.btn_convert)
        row_btns.addWidget(self.btn_cancel)
        lay_conv.addLayout(row_btns)
        right_lay.addWidget(grp_conv)

        right_lay.addStretch()
        right_scroll.setWidget(right_widget)
        splitter.addWidget(right_scroll)

        # Splitter proportions: 250px | flex | 350px
        splitter.setSizes([260, 440, 380])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)

        self.statusBar().showMessage("Ready")

    # ── Empty slides placeholder ──────────────────────────────────

    def _show_empty_slides(self):
        self.slide_list.clear()
        item = QListWidgetItem()
        item.setSizeHint(QSize(300, 200))
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setAlignment(Qt.AlignCenter)
        icon = QLabel("Slides will appear here")
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("color: #585b70; font-size: 14px;")
        sub = QLabel("Select a project or import a PPTX")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet("color: #45475a; font-size: 12px;")
        lay.addWidget(icon)
        lay.addWidget(sub)
        self.slide_list.addItem(item)
        self.slide_list.setItemWidget(item, w)

    # ── Project list ──────────────────────────────────────────────

    def _refresh_project_list(self, keep_selection=True):
        """Rebuild the project list from the projects directory."""
        prev_name = self._current_project_name if keep_selection else None
        self.project_list.clear()

        folders = sorted(
            p for p in Path(self.projects_dir).iterdir()
            if p.is_dir() and (p / PROJECT_JSON).exists()
        )

        selected_row = -1
        for i, folder in enumerate(folders):
            name = folder.name
            card = ProjectCard(name, str(folder))
            card.btn_folder.clicked.connect(
                lambda checked, d=str(folder): self._open_folder(d)
            )
            item = QListWidgetItem()
            item.setSizeHint(card.sizeHint())
            item.setData(Qt.UserRole, name)
            self.project_list.addItem(item)
            self.project_list.setItemWidget(item, card)
            if name == prev_name:
                selected_row = i

        self.lbl_proj_count.setText(f"{len(folders)} project(s)")

        if selected_row >= 0:
            self.project_list.blockSignals(True)
            self.project_list.setCurrentRow(selected_row)
            self.project_list.blockSignals(False)

    def _on_project_selected(self, row: int):
        """When a project row is clicked (single click), load it."""
        if row < 0:
            self.btn_rename_proj.setEnabled(False)
            self.btn_delete_proj.setEnabled(False)
            return
        self.btn_rename_proj.setEnabled(True)
        self.btn_delete_proj.setEnabled(True)
        self._load_selected_project()

    def _load_selected_project(self):
        """Load the currently selected project from the list."""
        row = self.project_list.currentRow()
        if row < 0:
            return
        item = self.project_list.item(row)
        name = item.data(Qt.UserRole)
        proj_dir = os.path.join(self.projects_dir, name)
        pj_file = os.path.join(proj_dir, PROJECT_JSON)
        if not os.path.isfile(pj_file):
            return
        self._load_project_from(name, proj_dir)

    def _load_project_from(self, name: str, proj_dir: str):
        """Load a project by name and directory path."""
        pj_file = os.path.join(proj_dir, PROJECT_JSON)
        try:
            with open(pj_file, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            QMessageBox.warning(self, "Load Error", f"Failed to load project:\n{exc}")
            return

        self._current_project_name = name
        self._current_project_dir = proj_dir
        self._current_pptx = data.get("pptx_path", "")

        # Apply settings to UI
        self.le_music.setText(data.get("music_path", ""))
        self.le_output.setText(data.get("output_path", ""))
        self.combo_res.setCurrentText(
            data.get("resolution_key", "1920 x 1080  (Full HD)")
        )
        custom = data.get("custom_resolution")
        if custom:
            self.spin_w.setValue(custom[0])
            self.spin_h.setValue(custom[1])
        self.spin_dur.setValue(data.get("seconds_per_slide", 5.0))
        self.combo_fps.setCurrentText(str(data.get("fps", 30)))
        self.combo_quality.setCurrentIndex(data.get("quality_index", 1))
        self.spin_volume.setValue(data.get("volume", 100))
        self._slide_count = data.get("slide_count", 0)

        # Load cached thumbnails
        self._load_cached_thumbnails()

        self.statusBar().showMessage(f"Loaded project: {name}")
        self.lbl_status.setText(f"Project: {name}")
        self._auto_save()

    def _load_cached_thumbnails(self):
        """Load thumbnails from the project's slides/ directory."""
        self.slide_list.clear()
        if not self._current_project_dir:
            self._show_empty_slides()
            self.lbl_slide_count.setText("")
            return

        slides_dir = os.path.join(self._current_project_dir, SLIDES_DIR)
        thumbs = sorted(Path(slides_dir).glob("*.png")) if os.path.isdir(slides_dir) else []

        if not thumbs:
            item = QListWidgetItem()
            item.setSizeHint(QSize(300, 120))
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setAlignment(Qt.AlignCenter)
            lbl = QLabel("No slide thumbnails yet.\nImport a PPTX to generate them.")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("color: #585b70; font-size: 13px;")
            lay.addWidget(lbl)
            self.slide_list.addItem(item)
            self.slide_list.setItemWidget(item, w)
            self.lbl_slide_count.setText("")
            return

        for i, thumb_path in enumerate(thumbs):
            item = QListWidgetItem()
            item.setSizeHint(QSize(196, 148))
            widget = SlideThumbWidget(i, str(thumb_path))
            self.slide_list.addItem(item)
            self.slide_list.setItemWidget(item, widget)

        self.lbl_slide_count.setText(f"{len(thumbs)} slide(s)")

    def _on_project_context_menu(self, pos):
        """Right-click context menu on the project list."""
        row = self.project_list.currentRow()
        menu = QMenu(self)
        if row >= 0:
            menu.addAction("Open", self._load_selected_project)
            menu.addSeparator()
            menu.addAction("Rename", self._rename_project)
            menu.addAction("Duplicate", self._duplicate_project)
            menu.addSeparator()
            menu.addAction("Delete", self._delete_project)
            menu.addSeparator()
            menu.addAction("Open in File Manager", self._open_project_folder)
        else:
            menu.addAction("New Project", self._new_project)
        menu.exec_(self.project_list.mapToGlobal(pos))

    # ── Project CRUD ──────────────────────────────────────────────

    def _new_project(self):
        name, ok = QInputDialog.getText(
            self, "New Project", "Project name:",
            text="",
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        # Sanitize
        safe = "".join(c for c in name if c.isalnum() or c in " -_.")
        safe = safe.strip(". ")
        if not safe:
            QMessageBox.warning(self, "Invalid Name", "Project name cannot be empty.")
            return
        proj_dir = os.path.join(self.projects_dir, safe)
        if os.path.exists(proj_dir):
            QMessageBox.warning(self, "Exists",
                                f"A project named '{safe}' already exists.")
            return

        os.makedirs(os.path.join(proj_dir, SLIDES_DIR), exist_ok=True)
        # Create empty project.json
        with open(os.path.join(proj_dir, PROJECT_JSON), "w", encoding="utf-8") as f:
            json.dump({
                "version": 3,
                "pptx_path": "",
                "music_path": "",
                "output_path": "",
                "resolution_key": "1920 x 1080  (Full HD)",
                "custom_resolution": None,
                "seconds_per_slide": 5.0,
                "fps": 30,
                "quality_index": 1,
                "volume": 100,
                "slide_count": 0,
            }, f, indent=2, ensure_ascii=False)

        self._refresh_project_list()
        # Select the new project
        for i in range(self.project_list.count()):
            if self.project_list.item(i).data(Qt.UserRole) == safe:
                self.project_list.setCurrentRow(i)
                break
        self.statusBar().showMessage(f"Created project: {safe}")

    def _rename_project(self):
        if not self._current_project_name:
            return
        new_name, ok = QInputDialog.getText(
            self, "Rename Project", "New name:",
            text=self._current_project_name,
        )
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()
        safe = "".join(c for c in new_name if c.isalnum() or c in " -_.")
        safe = safe.strip(". ")
        if not safe or safe == self._current_project_name:
            return

        old_dir = os.path.join(self.projects_dir, self._current_project_name)
        new_dir = os.path.join(self.projects_dir, safe)
        if os.path.exists(new_dir):
            QMessageBox.warning(self, "Exists",
                                f"A project named '{safe}' already exists.")
            return

        try:
            shutil.move(old_dir, new_dir)
            self._current_project_name = safe
            self._current_project_dir = new_dir
            self._refresh_project_list()
            self.statusBar().showMessage(f"Renamed to: {safe}")
        except Exception as exc:
            QMessageBox.critical(self, "Rename Error", str(exc))

    def _duplicate_project(self):
        if not self._current_project_name:
            return
        dup_name, ok = QInputDialog.getText(
            self, "Duplicate Project", "Name for the copy:",
            text=f"{self._current_project_name} (copy)",
        )
        if not ok or not dup_name.strip():
            return
        dup_name = dup_name.strip()
        safe = "".join(c for c in dup_name if c.isalnum() or c in " -_.")
        safe = safe.strip(". ")
        if not safe:
            return
        old_dir = os.path.join(self.projects_dir, self._current_project_name)
        new_dir = os.path.join(self.projects_dir, safe)
        if os.path.exists(new_dir):
            QMessageBox.warning(self, "Exists",
                                f"A project named '{safe}' already exists.")
            return
        try:
            shutil.copytree(old_dir, new_dir)
            self._refresh_project_list()
            for i in range(self.project_list.count()):
                if self.project_list.item(i).data(Qt.UserRole) == safe:
                    self.project_list.setCurrentRow(i)
                    break
            self.statusBar().showMessage(f"Duplicated as: {safe}")
        except Exception as exc:
            QMessageBox.critical(self, "Duplicate Error", str(exc))

    def _delete_project(self):
        if not self._current_project_name:
            return
        reply = QMessageBox.question(
            self, "Delete Project",
            f"Delete project '{self._current_project_name}'?\n\n"
            f"This will permanently remove the project folder and all its files.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        try:
            shutil.rmtree(self._current_project_dir)
            self._current_project_name = None
            self._current_project_dir = None
            self._current_pptx = ""
            self._slide_count = 0
            self._refresh_project_list()
            self._show_empty_slides()
            self.lbl_status.setText("Project deleted.")
            self.le_music.clear()
            self.le_output.clear()
            self.statusBar().showMessage("Project deleted")
        except Exception as exc:
            QMessageBox.critical(self, "Delete Error", str(exc))

    def _open_project_folder(self):
        """Open the current project's folder (or the projects root) in file manager."""
        target = (
            self._current_project_dir
            if self._current_project_dir
            else self.projects_dir
        )
        self._open_folder(target)

    def _open_folder(self, path: str):
        """Open a folder in the system file manager."""
        import subprocess
        import sys
        path = os.path.abspath(path)
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])

    def _change_projects_dir(self):
        path = QFileDialog.getExistingDirectory(
            self, "Select Projects Directory", self.projects_dir,
        )
        if path:
            self.projects_dir = path
            self.lbl_proj_dir.setText(f"  {path}")
            self._current_project_name = None
            self._current_project_dir = None
            self._refresh_project_list(keep_selection=False)
            self._show_empty_slides()
            self.statusBar().showMessage(f"Projects directory: {path}")

    # ── Auto-save project.json ────────────────────────────────────

    def _auto_save(self):
        """Write current UI state to project.json (called after setting changes)."""
        if not self._current_project_dir:
            return
        pj_file = os.path.join(self._current_project_dir, PROJECT_JSON)
        data = {
            "version": 3,
            "pptx_path": self._current_pptx,
            "music_path": self.le_music.text().strip(),
            "output_path": self.le_output.text().strip(),
            "resolution_key": self.combo_res.currentText(),
            "custom_resolution": (
                [self.spin_w.value(), self.spin_h.value()]
                if self.combo_res.currentIndex() == len(RESOLUTION_PRESETS) - 1
                else None
            ),
            "seconds_per_slide": self.spin_dur.value(),
            "fps": int(self.combo_fps.currentText()),
            "quality_index": self.combo_quality.currentIndex(),
            "volume": self.spin_volume.value(),
            "slide_count": self._slide_count,
        }
        try:
            with open(pj_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    # ── Drag & Drop ───────────────────────────────────────────────

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(u.toLocalFile().lower().endswith((".pptx", ".ppt")) for u in urls):
                event.acceptProposedAction()
                return
        super().dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith((".pptx", ".ppt")):
                self._import_pptx_to_project(path)
                event.acceptProposedAction()
                return
        super().dropEvent(event)

    # ── PPTX import ───────────────────────────────────────────────

    def _browse_pptx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Presentation", "",
            "PowerPoint Files (*.pptx *.ppt);;All Files (*)",
        )
        if path:
            self._import_pptx_to_project(path)

    def _import_pptx_to_project(self, path: str):
        """Import a PPTX file into the current project (or create one)."""
        # If no project selected, create one from the PPTX filename
        if not self._current_project_dir:
            stem = Path(path).stem
            safe = "".join(c for c in stem if c.isalnum() or c in " -_.")
            safe = safe.strip(". ") or "Untitled"
            proj_dir = os.path.join(self.projects_dir, safe)
            # If already exists, just use it
            if not os.path.exists(proj_dir):
                os.makedirs(os.path.join(proj_dir, SLIDES_DIR), exist_ok=True)
                with open(os.path.join(proj_dir, PROJECT_JSON), "w", encoding="utf-8") as f:
                    json.dump({
                        "version": 3, "pptx_path": "", "music_path": "",
                        "output_path": "", "resolution_key": "1920 x 1080  (Full HD)",
                        "custom_resolution": None, "seconds_per_slide": 5.0,
                        "fps": 30, "quality_index": 1, "volume": 100,
                        "slide_count": 0,
                    }, f, indent=2, ensure_ascii=False)
                self._refresh_project_list()
                for i in range(self.project_list.count()):
                    if self.project_list.item(i).data(Qt.UserRole) == safe:
                        self.project_list.setCurrentRow(i)
                        break
            else:
                self._load_project_from(safe, proj_dir)

        self._current_pptx = path
        # Auto-set output path
        stem = Path(path).stem
        parent = Path(path).parent
        self.le_output.setText(str(parent / f"{stem}.mp4"))

        # Generate thumbnails into the project's slides/ folder
        slides_dir = os.path.join(self._current_project_dir, SLIDES_DIR)
        self.btn_import.setEnabled(False)
        self.btn_import.setText("Loading...")
        self.slide_list.clear()
        self.statusBar().showMessage(f"Generating thumbnails for {os.path.basename(path)}...")

        self.thumb_worker = ThumbnailWorker(path, slides_dir)
        self.thumb_worker.thumbnail_ready.connect(self._on_thumbnail_ready)
        self.thumb_worker.all_done.connect(self._on_thumbnails_done)
        self.thumb_worker.error.connect(self._on_thumb_error)
        self.thumb_worker.progress.connect(lambda pct, msg: self.statusBar().showMessage(msg))
        self.thumb_worker.start()

    def _on_thumbnail_ready(self, index: int, path: str):
        item = QListWidgetItem()
        item.setSizeHint(QSize(196, 148))
        widget = SlideThumbWidget(index, path)
        self.slide_list.addItem(item)
        self.slide_list.setItemWidget(item, widget)

    def _on_thumbnails_done(self, count: int):
        self._slide_count = count
        self.btn_import.setEnabled(True)
        self.btn_import.setText("Import PPTX...")
        self.lbl_slide_count.setText(f"{count} slide(s)")
        self.statusBar().showMessage(f"Loaded {count} slides")
        self._auto_save()
        self._refresh_project_list()  # update cards with new thumbnails
        self.thumb_worker = None

    def _on_thumb_error(self, msg: str):
        self.btn_import.setEnabled(True)
        self.btn_import.setText("Import PPTX...")
        self.statusBar().showMessage("Thumbnail generation failed")
        QMessageBox.warning(
            self, "Import Error",
            f"Failed to generate thumbnails:\n\n{msg}\n\n"
            "You can still convert, but previews won't update.",
        )
        self.thumb_worker = None

    # ── Music / Output / Settings ─────────────────────────────────

    def _browse_music(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Music File", "",
            "Audio Files (*.mp3 *.wav *.ogg *.flac *.aac *.m4a);;All Files (*)",
        )
        if path:
            self.le_music.setText(path)
            self._auto_save()

    def _clear_music(self):
        self.le_music.clear()
        self._auto_save()

    def _browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Video As", "", "MP4 Video (*.mp4);;All Files (*)",
        )
        if path:
            if not path.lower().endswith(".mp4"):
                path += ".mp4"
            self.le_output.setText(path)
            self._auto_save()

    def _on_resolution_changed(self, idx):
        self.custom_widget.setVisible(idx == len(RESOLUTION_PRESETS) - 1)
        self._auto_save()

    def _get_resolution(self):
        key = self.combo_res.currentText()
        val = RESOLUTION_PRESETS.get(key)
        if val:
            return val
        return (self.spin_w.value(), self.spin_h.value())

    def _get_quality_value(self):
        return [5, 18, 28][self.combo_quality.currentIndex()]

    # ── Conversion ────────────────────────────────────────────────

    def _start_conversion(self):
        pptx = self._current_pptx.strip()
        if not pptx or not os.path.isfile(pptx):
            QMessageBox.warning(self, "Missing Input",
                                "Please import a valid .pptx file first.")
            return
        output = self.le_output.text().strip()
        if not output:
            QMessageBox.warning(self, "Missing Output",
                                "Please specify an output .mp4 path.")
            return
        music = self.le_music.text().strip() or None
        if music and not os.path.isfile(music):
            QMessageBox.warning(self, "Invalid Music",
                                f"Music file not found:\n{music}")
            return
        if not shutil.which("libreoffice"):
            QMessageBox.critical(self, "Missing Dependency",
                                 "LibreOffice is required but not found.")
            return
        if not shutil.which("ffmpeg"):
            QMessageBox.critical(self, "Missing Dependency",
                                 "ffmpeg is required but not found.")
            return

        res = self._get_resolution()
        fps = int(self.combo_fps.currentText())
        dur = self.spin_dur.value()
        quality = self._get_quality_value()
        self._volume = self.spin_volume.value()

        self.progress_bar.setValue(0)
        self.btn_convert.setEnabled(False)
        self.btn_cancel.setEnabled(True)

        self.worker = ConversionWorker(
            pptx_path=pptx, output_path=output, music_path=music,
            seconds_per_slide=dur, fps=fps, resolution=res, quality=quality,
        )
        self._output_path = output
        self.worker.progress.connect(self._on_progress)
        self.worker.status_message.connect(self._on_status)
        self.worker.finished_ok.connect(self._on_success)
        self.worker.finished_error.connect(self._on_error)
        self.worker.start()

    def _cancel_conversion(self):
        if self.worker:
            self.worker.cancel()
        self.lbl_status.setText("Cancelling...")

    def _on_progress(self, value):
        self.progress_bar.setValue(value)

    def _on_status(self, msg):
        self.lbl_status.setText(msg)
        self.statusBar().showMessage(msg)

    def _on_success(self, path):
        if getattr(self, '_volume', 100) != 100 and self.le_music.text().strip():
            self.lbl_status.setText(self.lbl_status.text() + "\nAdjusting volume...")
            self._adjust_volume(path, self._volume)
        self.btn_convert.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress_bar.setValue(100)
        self.lbl_status.setText(f"Conversion complete!\n{path}")
        self.statusBar().showMessage("Done!")
        QMessageBox.information(self, "Success", f"Video saved to:\n{path}")
        self.worker = None

    def _adjust_volume(self, video_path, volume_pct):
        factor = volume_pct / 100.0
        tmp = video_path + ".tmp.mp4"
        cmd = [
            "ffmpeg", "-y", "-i", video_path,
            "-af", f"volume={factor}",
            "-c:v", "copy", "-c:a", "aac", "-b:a", "192k", tmp,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if result.returncode == 0 and os.path.isfile(tmp):
            shutil.move(tmp, video_path)
        elif os.path.exists(tmp):
            os.remove(tmp)

    def _on_error(self, msg):
        self.btn_convert.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress_bar.setValue(0)
        self.lbl_status.setText(f"Error: {msg}")
        self.statusBar().showMessage("Conversion failed")
        QMessageBox.critical(self, "Conversion Error",
                             f"An error occurred:\n\n{msg}")
        self.worker = None

    # ── About ─────────────────────────────────────────────────────

    def _show_about(self):
        QMessageBox.about(
            self, "About",
            "<h3>PPTX to MP4 Converter v3</h3>"
            "<p>Convert PowerPoint to MP4 with project management.</p>"
            "<p>Projects are stored in: <code>"
            + self.projects_dir + "</code></p>"
        )

    # ── Window settings ───────────────────────────────────────────

    def _restore_window_settings(self):
        try:
            if os.path.isfile(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r") as f:
                    s = json.load(f)
                self.resize(s.get("width", 1160), s.get("height", 780))
                if s.get("maximized"):
                    self.showMaximized()
                if s.get("projects_dir"):
                    pdir = s["projects_dir"]
                    if os.path.isdir(pdir):
                        self.projects_dir = pdir
                        self.lbl_proj_dir.setText(f"  {pdir}")
                        self._refresh_project_list(keep_selection=False)
        except Exception:
            pass

    def _save_window_settings(self):
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump({
                    "width": self.width(), "height": self.height(),
                    "maximized": self.isMaximized(),
                    "projects_dir": self.projects_dir,
                }, f)
        except Exception:
            pass

    # ── Cleanup ───────────────────────────────────────────────────

    def closeEvent(self, event):
        if self.thumb_worker and self.thumb_worker.isRunning():
            self.thumb_worker.cancel()
            self.thumb_worker.wait(3000)
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "Confirm Exit",
                "Conversion is still running. Cancel and exit?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.worker.cancel()
                self.worker.wait(3000)
                event.accept()
            else:
                event.ignore()
            return
        self._save_window_settings()
        event.accept()


# ═══════════════════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#1e1e2e"))
    palette.setColor(QPalette.WindowText, QColor("#cdd6f4"))
    palette.setColor(QPalette.Base, QColor("#313244"))
    palette.setColor(QPalette.AlternateBase, QColor("#45475a"))
    palette.setColor(QPalette.ToolTipBase, QColor("#45475a"))
    palette.setColor(QPalette.ToolTipText, QColor("#cdd6f4"))
    palette.setColor(QPalette.Text, QColor("#cdd6f4"))
    palette.setColor(QPalette.Button, QColor("#45475a"))
    palette.setColor(QPalette.ButtonText, QColor("#cdd6f4"))
    palette.setColor(QPalette.BrightText, QColor("#f38ba8"))
    palette.setColor(QPalette.Highlight, QColor("#89b4fa"))
    palette.setColor(QPalette.HighlightedText, QColor("#1e1e2e"))
    app.setPalette(palette)
    app.setStyleSheet(STYLESHEET)

    window = PPTXToMP4Converter()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()