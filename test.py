import sys


from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtWidgets import QLabel, QLineEdit


class Trick(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(300, 300, 320, 60)
        self.setWindowTitle('Вычисление выражений')

        self.untr = QLineEdit(self)
        self.untr.setGeometry(10, 20, 125, 30)

        self.untr_label = QLabel(self)
        self.untr_label.setText('Выражение:')
        self.untr_label.move(10, 5)

        self.tr = QLineEdit(self)
        self.tr.setGeometry(175, 20, 125, 30)
        self.untr_label = QLabel(self)
        self.untr_label.setText('Результат:')
        self.untr_label.move(175, 5)

        self.btn = QPushButton('->', self)
        self.btn.setGeometry(140, 20, 30, 30)
        self.btn.clicked.connect(self.tricked)

    def tricked(self):
        self.tr.setText(eval(self.untr.text()))
        self.btn.setText('->')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Trick()
    ex.show()
    sys.exit(app.exec())