import random
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QMainWindow, QApplication, QPushButton, QWidget

class MainWindow(QMainWindow):
	def __init__(self, appname):
		self.appname = appname
		super().__init__()
	
		self.setWindowTitle(appname)
		button = QPushButton("x")
		self.setCentralWidget(button)


	def get_appname(self):
		return self.appname


if __name__ == "__main__":
	app = QApplication([])

	window = QMainWindow()
	window.show()

	app.exec()
