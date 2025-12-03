import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QPushButton

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("示例应用程序")
        self.setGeometry(100, 100, 400, 300)

        # 确认退出按钮
        self.exit_button = QPushButton("关闭窗口", self)
        self.exit_button.setGeometry(150, 125, 100, 30)
        self.exit_button.clicked.connect(self.show_exit_confirmation)

    def show_exit_confirmation(self):
        # 创建确认对话框
        reply = QMessageBox.question(self, "确认", "您确定要关闭窗口吗？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            QApplication.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())