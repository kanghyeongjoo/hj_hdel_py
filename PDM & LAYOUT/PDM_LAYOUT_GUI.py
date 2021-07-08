import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import *
from PDM_LAYOUT_DATA import get_proj_data

class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PLM & LAYOUT DATA 비교')
        self.setWindowIcon(QIcon("D:\OneDrive - 현대그룹\logo.png"))
        self.setGeometry(80, 80, 1000, 800)
        # self.move(50, 50)
        # self.resize(1000, 800)
        #
        self.prg_bar = QProgressBar()
        self.prg_bar.setValue(0)
        self.prg_bar.setAlignment(Qt.AlignCenter)

        self.run_timer = QTimer()
        self.run_timer.setInterval(1000)
        self.run_timer.timeout.connect(self.run_time)

        # c_day = QDate.currentDate()
        # c_time = QTime.currentTime()
        c_date_time = QDateTime.currentDateTime()
        state = "Ready"
        self.stt_bar = QtWidgets.QStatusBar()
        self.setStatusBar(self.stt_bar)
        self.stt_bar.showMessage("상태 : "+ state + " "*20 + c_date_time.toString("yyyy-MM-dd hh:mm:ss"))

        exitAct = QAction(QIcon("D:\OneDrive - 현대그룹\exit.png"), "Exit", self)
        exitAct.triggered.connect(qApp.quit)
        saveAct = QAction(QIcon("D:\OneDrive - 현대그룹\save.png"), "Save", self)
        saveAct.triggered.connect(qApp.quit)
        self.exit_tool_Bar = self.addToolBar("Exit")
        self.exit_tool_Bar.addAction(exitAct)
        self.save_tool_bar = self.addToolBar("Save")
        self.save_tool_bar.addAction(saveAct)


        self.proj_label = QLabel("현장번호")
        self.proj_label.setAlignment(Qt.AlignCenter)
        self.proj_label_font = self.proj_label.font()
        # proj_label_font.setPointSize(10)
        self.proj_label_font.setBold(True)
        self.proj_label.setFont(self.proj_label_font)

        self.proj_no_input = QLineEdit()
        self.proj_no_input.resize(120, 25)
        self.proj_no_input.returnPressed.connect(self.proj_run)

        self.run_btn = QPushButton("비교 실행")
        self.run_btn.resize(100, 30)
        self.run_btn.clicked.connect(self.proj_run)

        self.tabs = QTabWidget()
        self.total_tab = QWidget()
        self.match_tab = QWidget()
        self.chk_tab = QWidget()
        self.tabs.addTab(self.total_tab, 'All data')
        self.tabs.addTab(self.match_tab, 'Match data')
        self.tabs.addTab(self.chk_tab, 'Check data')
        tap_h = self.height()-150
        self.tabs.resize(self.width() - 40, tap_h)

        layout_h = QHBoxLayout()
        layout_h.setContentsMargins(20, 0, 0, 0)
        layout_h.addWidget(self.proj_label)
        layout_h.addWidget(self.proj_no_input)
        layout_h.addSpacing(30)
        layout_h.addWidget(self.run_btn)
        layout_h.addStretch(1)
        layout_v = QVBoxLayout()
        layout_v.setContentsMargins(0, 30, 0, 0)
        layout_v.addLayout(layout_h)
        layout_v.addSpacing(10)
        layout_v.addWidget(self.tabs)
        widget = QWidget()
        widget.setLayout(layout_v)
        self.setCentralWidget(widget)

        self.total_table = QTableWidget(self)
        self.total_tab.layout = QVBoxLayout()
        self.total_tab.layout.setContentsMargins(5, 5, 5, 5)
        self.total_tab.layout.addWidget(self.total_table)
        self.total_tab.setLayout(self.total_tab.layout)

        self.match_table = QTableWidget(self)
        self.match_tab.layout = QVBoxLayout()
        self.match_tab.layout.setContentsMargins(5, 5, 5, 5)
        self.match_tab.layout.addWidget(self.match_table)
        self.match_tab.setLayout(self.match_tab.layout)

        self.chk_table = QTableWidget(self)
        self.chk_tab.layout = QVBoxLayout()
        self.chk_tab.layout.setContentsMargins(10, 5, 0, 0)
        self.chk_tab.layout.addWidget(self.chk_table)
        self.chk_tab.setLayout(self.chk_tab.layout)

    def run_time(self):
        time = self.prg_bar.value()
        time += 5
        self.prg_bar.setValue(time)

        if time >= self.prg_bar.maximum():
            self.run_timer.stop()


    def proj_run(self):
        # self.stt_bar.addPermanentWidget(self.prg_bar)
        # self.run_timer.start()
        QApplication.setOverrideCursor(Qt.WaitCursor)

        proj_no = self.proj_no_input.text()
        self.total_data, self.match_data, self.chk_data, self.floor_data = get_proj_data(proj_no)

        total_table_col = list(self.total_data.index)
        total_table_row = list(self.total_data.columns)
        self.total_table.setColumnCount(len(total_table_col))
        self.total_table.setRowCount(len(total_table_row))
        self.total_table.setHorizontalHeaderLabels(total_table_col)

        match_table_col = list(self.match_data.index)
        match_table_row = list(self.match_data.columns)
        self.match_table.setColumnCount(len(match_table_col))
        self.match_table.setRowCount(len(match_table_row))
        self.match_table.setHorizontalHeaderLabels(match_table_col)

        for total_row_no in range(len(total_table_row)-1):
            for total_col_no in range(len(total_table_col)):
                self.total_table.setItem(total_row_no, total_col_no, QTableWidgetItem(str(self.total_data.iloc[total_col_no, total_row_no])))

        for match_row_no in range(len(match_table_row)):
            for match_col_no in range(len(match_table_col)):
                self.match_table.setItem(match_row_no, match_col_no, QTableWidgetItem(str(self.match_data.iloc[match_col_no, match_row_no])))

        if not self.chk_data.empty:
            chk_table_col = list(self.chk_data.index)
            chk_table_row = list(self.chk_data.columns)
            self.chk_table.setColumnCount(len(chk_table_col))
            self.chk_table.setRowCount(len(chk_table_row))
            self.chk_table.setHorizontalHeaderLabels(chk_table_col)
            for chk_row_no in range(len(chk_table_row)):
                for chk_col_no in range(len(chk_table_col)):
                    self.chk_table.setItem(chk_row_no, chk_col_no,
                                           QTableWidgetItem(str(self.chk_data.iloc[chk_col_no, chk_row_no])))
        QApplication.setOverrideCursor(Qt.ArrowCursor)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MyWindow = MyApp()
    MyWindow.show()
    sys.exit(app.exec_())