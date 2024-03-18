from db_handler import db_handler
import sys
from PyQt5.QtWidgets import (
	QApplication, QDialog, QMainWindow, QMessageBox, QTreeWidgetItem
)
from ui.main_window import Ui_MainWindow

class Window(QMainWindow, Ui_MainWindow):
	def __init__(self, parent=None):
		super().__init__(parent)
		self.setupUi(self)

	def populate_tree(self, data):
		items = []
		for key, values in data.items():
			item = QTreeWidgetItem([key])
			for value in values:
				child = QTreeWidgetItem(value)
				item.addChild(child)
			items.append(item)
		self.treeWidget.clear()
		self.treeWidget.insertTopLevelItems(0, items)

if __name__ == "__main__":
	app = QApplication(sys.argv)
	win = Window()
	data = db.build_tree_data()
	win.populate_tree(data)
	win.show()
	sys.exit(app.exec())