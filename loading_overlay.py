from PyQt5.QtWidgets import QWidget, QProgressBar, QLabel, QVBoxLayout
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPainter, QColor

class LoadingOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Make the overlay cover the entire parent widget
        self.setFixedSize(parent.size())
        self.hide()
        
        # Semi-transparent background
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("""
            QWidget {
                background-color: rgba(0, 0, 0, 100);
            }
            QLabel {
                color: white;
                font-size: 14px;
                background-color: transparent;
            }
            QProgressBar {
                border: 2px solid #4CAF50;
                border-radius: 5px;
                text-align: center;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        """)
        
        # Create layout
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        
        # Add loading label
        self.label = QLabel("Processing...")
        layout.addWidget(self.label, alignment=Qt.AlignCenter)
        
        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedSize(300, 25)
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar, alignment=Qt.AlignCenter)
        
        self.setLayout(layout)
        
        # Timer for simulating progress
        self.timer = QTimer()
        self.timer.timeout.connect(self._update_progress)
        self.current_progress = 0
    
    def start_loading(self, message="Processing..."):
        self.label.setText(message)
        self.progress_bar.setValue(0)
        self.current_progress = 0
        self.show()
        self.timer.start(50)  # Update every 50ms
        
    def stop_loading(self):
        self.timer.stop()
        self.hide()
    
    def _update_progress(self):
        if self.current_progress < 95:  # Cap at 95% until operation completes
            self.current_progress += 1
            self.progress_bar.setValue(self.current_progress)
    
    def set_progress(self, value):
        self.current_progress = value
        self.progress_bar.setValue(value)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100))
        super().paintEvent(event)
