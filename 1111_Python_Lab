from PyQt5.QtWidgets import *
from PyQt5.QtGui import QFont, QKeySequence
from PyQt5.QtCore import Qt
import math

# project modules needed
import nidmm
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from zaber.serial import BinarySerial, BinaryDevice


class NewWindow(QDialog):
    def __init__(self, number: int):
        super().__init__()

        self.setWindowTitle("Picture")
        self.setGeometry(660, 240, 600, 600)

        self.figure = Figure()

        self.canvas = FigureCanvas(self.figure)

        self.toolbar = NavigationToolbar(self.canvas, self)

        layout = QVBoxLayout(self)

        layout.addWidget(self.canvas)
        layout.addWidget(self.toolbar)

        with BinarySerial("COM7") as port:
            deg = 4267
            seg = 360 / (number + 1)
            device = BinaryDevice(port, 1)
            reply = device.home()
            if self.check_command_succeeded(reply):
                print("Device homed.")
            else:
                print("Device home failed.")
                exit(1)

            lst_degree = []
            lst_measure = []
            for N in range(number):
                reply = device.move_rel(seg * deg)
                lst_degree.append(seg * N)
                with nidmm.Session("Dev1") as session:
                    lst_measure.append(session.read())

        self.plot(lst_degree, lst_measure)

    def plot(self, lst_x: list, lst_y: list):
        lst_x = [round(elem, 2) for elem in lst_x]
        max_y, min_y = math.ceil(max(lst_y)), math.floor(min(lst_y))
        # create a layout containing a subplot
        # subplot(row, col, index)
        xy_graph = self.figure.add_subplot(2, 1, 1)

        xy_graph.set_xticks(lst_x[1])
        xy_graph.set_xticklabels(lst_x)

        xy_graph_ylbls = []
        if max_y - min_y <= 4:
            xy_graph.set_yticks(0.25)
            m, M = min_y * 100, max_y * 100

            for elem in range(m, M + 1, 25):
                xy_graph_ylbls.append(elem / 100)
        else:
            xy_graph.set_yticks(1)
            for elem in range(m, M + 1):
                xy_graph_ylbls.append(elem)

        xy_graph.set_yticklabels(xy_graph_ylbls)

        xy_graph.set_title("XY Chart: ° vs Mea")
        xy_graph.set_xlabel("degree (°)")
        xy_graph.set_ylabel("Intensity")
        # draw on the figure
        xy_graph.plot(lst_x, lst_y)

        # plot a radar chart
        radar_chart = self.figure.add_subplot(2, 1, 2)

        radar_chart.set_xticks(lst_x[1])
        radar_chart.set_xticklabels(lst_x)

        radar_chart.set_ylim(0, max_y)

        radar_chart.set_title("Radar Chart")
        # renew the picture
        self.canvas.draw()

    def check_command_succeeded(self, reply):
        if reply.command_number == 255:  # 255 is the binary error response code.
            QMessageBox.warning(
                self, "Warning", "Danger! Command rejected.Error code: " + str(reply.data))
            return False
        else:
            return True


class MyWindow(QMainWindow):
    __cnt = 0

    def __init__(self):
        super().__init__()

        # build window
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.setWindowTitle("Malus Experiment Test")

        # select main layout
        layout = QVBoxLayout(central_widget)

        # put sub layout in main layout
        subLayout_1 = QHBoxLayout()
        layout.addLayout(subLayout_1)

        # sub lay out widgets
        lbl_N_Datas = QLabel()
        lbl_N_Datas.setText("Data N =")
        lbl_N_Datas.setFont(QFont('Consolas', 20))
        subLayout_1.addWidget(lbl_N_Datas)

        self.txt_N_Datas = QLineEdit()
        self.txt_N_Datas.setFixedWidth(50)
        self.txt_N_Datas.setFixedHeight(40)
        self.txt_N_Datas.setMaxLength(3)
        self.txt_N_Datas.setFont(QFont('Consolas', 16))
        subLayout_1.addWidget(self.txt_N_Datas)

        btn_submit = QPushButton()
        btn_submit.setText("SUBMIT")
        btn_submit.setFont(QFont('Consolas', 20))
        btn_submit.setFixedHeight(40)
        btn_submit.setFixedWidth(100)
        subLayout_1.addWidget(btn_submit)

        # button clicked
        btn_submit.clicked.connect(self.btn_submit_clicked)

        shortcut_btn_submit = QShortcut(QKeySequence(Qt.Key_Return), self)
        shortcut_btn_submit.activated.connect(self.btn_submit_clicked)

    def btn_submit_clicked(self):
        number = self.txt_N_Datas.text()
        if not number.isdigit():
            mes = QMessageBox()
            tempFont = QFont('Consolas', 16)
            tempFont.setBold(True)
            mes.setWindowTitle("Warning")
            mes.setText("Wrong Input!")
            mes.setFont(tempFont)
            mes.setStyleSheet(
                "QMessageBox QLabel { color: red; } QMessageBox QPushButton { color: black }")
            mes.setStandardButtons(QMessageBox.Ok)
            mes.exec_()
        elif int(number) > 100 or int(number) <= 0:
            mes = QMessageBox()
            tempFont = QFont('Consolas', 16)
            mes.setText("We'll set N = 100 for simplicity!")
            mes.setWindowTitle("Notice")
            mes.setStyleSheet(
                "QMessageBox QLabel { color: red; } QMessageBox QPushButton { color: black }")
            mes.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            result = mes.exec_()
            if result == QMessageBox.Yes:
                pass
            else:
                self.__cnt += 1
                if self.__cnt != 3:
                    QMessageBox.warning(self, "Warning",
                                        "Please re-input the N, and 0 < N <= 100!")
                else:
                    QMessageBox.warning(self, "Fuck-you", "FUCK-YOU")
                    exit()
        else:
            self.hide()
            N = int(self.txt_N_Datas.text())
            newWindow = NewWindow(N)
            newWindow.exec_()
            self.show()


if __name__ == '__main__':
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()
