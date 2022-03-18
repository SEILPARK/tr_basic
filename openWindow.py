#-*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QApplication, QWidget

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('My First Application')
        self.move(300, 300)
        self.resize(400, 200)
        self.show()


if __name__ == '__main__': # 프로그램의 시작점일 때만 아래 코드 실행
    app = QApplication(sys.argv) # QApplication 객체 생성.
    print("안녕")
    ex = MyApp()
    sys.exit(app.exec_()) # 이벤트 루프 수행함으로써 프로그램이 종료되지 않음.