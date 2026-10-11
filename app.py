import sys
import os
import tempfile
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                               QComboBox, QSpinBox, QFileDialog, QTextEdit, QMessageBox)
from PySide6.QtCore import QThread, Signal
from PIL import Image, ImageFilter
import moviepy.editor as mp

# --- Processing Thread ---
class VideoWorker(QThread):
    progress = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, input_path, output_path, input_type, duration):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.input_type = input_type
        self.duration = duration
        self.temp_dir = tempfile.mkdtemp()

    def resize_to_shorts(self, img_path):
        """Resizes image to 1080x1920 (9:16) with blurred background"""
        self.progress.emit(f"Processing image: {os.path.basename(img_path)}")
        target_w, target_h = 1080, 1920
        img = Image.open(img_path).convert("RGB")
        
        # Create blurred background
        bg = img.resize((target_w, target_h))
        bg = bg.filter(ImageFilter.GaussianBlur(30))
        
        # Resize original to fit within the frame
        img.thumbnail((target_w, target_h), Image.LANCZOS)
        offset = ((target_w - img.width) // 2, (target_h - img.height) // 2)
        bg.paste(img, offset)
        
        out_path = os.path.join(self.temp_dir, f"frame_{os.path.basename(img_path)}")
        bg.save(out_path, "JPEG", quality=95)
        return out_path

    def extract_pptx_images(self):
        self.progress.emit("Extracting PPTX slides...")
        try:
            from pptx import Presentation
            from pdf2image import convert_from_path
        except ImportError:
            return None

        # PPTX doesn't easily render directly to images, so we convert to PDF first using LibreOffice/Office logic
        # For simplicity, we use a workaround: Convert PPTX to PDF using subprocess if possible, 
        # OR ask user to provide PDF. Here we will try a direct python-pptx -> pdf2image approach.
        # Note: python-pptx cannot render to image. A common workaround is converting to PDF first.
        # We will assume the user has LibreOffice installed for pptx to pdf conversion in this example.
        
        try:
            import subprocess
            pdf_path = os.path.join(self.temp_dir, "temp.pdf")
            # Requires LibreOffice installed
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', self.temp_dir, self.input_path], check=True)
            generated_pdf = os.path.join(self.temp_dir, os.path.splitext(os.path.basename(self.input_path))[0] + ".pdf")
            
            images = convert_from_path(generated_pdf)
            img_paths = []
            for i, image in enumerate(images):
                path = os.path.join(self.temp_dir, f"slide_{i}.png")
                image.save(path, "PNG")
                img_paths.append(path)
            return img_paths
        except Exception as e:
            self.progress.emit(f"Error extracting PPTX: {str(e)} (LibreOffice required for PPTX)")
            return None

    def extract_html_images(self):
        self.progress.emit("Rendering HTML...")
        try:
            from html2image import Html2Image
            hti = Html2Image(output_path=self.temp_dir)
            # Render HTML to a 1080x1920 canvas
            hti.screenshot(url=self.input_path, save_as="html_render.png", size=(1080, 1920))
            return [os.path.join(self.temp_dir, "html_render.png")]
        except Exception as e:
            self.progress.emit(f"Error rendering HTML: {str(e)}")
            return None

    def run(self):
        try:
            if self.input_type == "HTML":
                raw_images = self.extract_html_images()
            else:
                raw_images = self.extract_pptx_images()

            if not raw_images:
                self.finished.emit(False, "Failed to extract images from the input file.")
                return

            processed_images = []
            for img in raw_images:
                processed_images.append(self.resize_to_shorts(img))

            self.progress.emit("Creating video clips...")
            clips = []
            for img_path in processed_images:
                clip = mp.ImageClip(img_path).set_duration(self.duration)
                clips.append(clip)

            self.progress.emit("Concatenating and rendering MP4...")
            final_clip = mp.concatenate_videoclips(clips, method="compose")
            final_clip.write_videofile(
                self.output_path, 
                fps=24, 
                codec='libx264', 
                audio=False, 
                verbose=False, 
                logger=None
            )

            self.finished.emit(True, f"Video successfully saved to:\n{self.output_path}")

        except Exception as e:
            self.finished.emit(False, f"An error occurred: {str(e)}")

# --- Main GUI ---
class ShortsMakerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("YouTube Shorts Maker")
        self.setMinimumSize(600, 400)
        self.worker = None

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Input Type
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Input Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["HTML", "PPTX"])
        self.type_combo.currentTextChanged.connect(self.update_file_filter)
        type_layout.addWidget(self.type_combo)
        layout.addLayout(type_layout)

        # File Selection
        file_layout = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select input file...")
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)

        # Duration
        dur_layout = QHBoxLayout()
        dur_layout.addWidget(QLabel("Seconds per slide/page:"))
        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(1, 60)
        self.duration_spin.setValue(3)
        dur_layout.addWidget(self.duration_spin)
        layout.addLayout(dur_layout)

        # Output Selection
        out_layout = QHBoxLayout()
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("Select output MP4 file...")
        out_browse_btn = QPushButton("Browse")
        out_browse_btn.clicked.connect(self.browse_output)
        out_layout.addWidget(self.out_edit)
        out_layout.addWidget(out_browse_btn)
        layout.addLayout(out_layout)

        # Generate Button
        self.gen_btn = QPushButton("Generate YouTube Short")
        self.gen_btn.setStyleSheet("background-color: #ff0000; color: white; font-weight: bold; padding: 10px;")
        self.gen_btn.clicked.connect(self.generate_video)
        layout.addWidget(self.gen_btn)

        # Log/Progress
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

    def update_file_filter(self, text):
        if text == "HTML":
            self._file_filter = "HTML Files (*.html *.htm)"
        else:
            self._file_filter = "PowerPoint Files (*.pptx)"

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Input File", "", self._file_filter)
        if file_path:
            self.file_edit.setText(file_path)

    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Output Video", "", "MP4 Files (*.mp4)")
        if file_path:
            if not file_path.endswith('.mp4'):
                file_path += '.mp4'
            self.out_edit.setText(file_path)

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def generate_video(self):
        input_path = self.file_edit.text()
        output_path = self.out_edit.text()
        input_type = self.type_combo.currentText()
        duration = self.duration_spin.value()

        if not input_path or not os.path.exists(input_path):
            QMessageBox.warning(self, "Error", "Please select a valid input file.")
            return
        if not output_path:
            QMessageBox.warning(self, "Error", "Please select an output path.")
            return

        self.gen_btn.setEnabled(False)
        self.log_text.clear()
        self.log("Starting video generation...")

        self.worker = VideoWorker(input_path, output_path, input_type, duration)
        self.worker.progress.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success, message):
        self.log(message)
        self.gen_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", "Video generation complete!")
        else:
            QMessageBox.critical(self, "Error", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ShortsMakerApp()
    window.show()
    sys.exit(app.exec())