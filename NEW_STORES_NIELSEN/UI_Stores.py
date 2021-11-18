from PyQt5 import QtWidgets as qtw
from PyQt5 import QtGui as qtg
from PyQt5 import QtCore as qtc

country_input = ''
excel_name_input = ''
sheet_num_input = ''
rows_num_input = ''


class MainWindow(qtw.QWidget):

    # Class constructor
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.setup_UI()

        self.leCountry.returnPressed.connect(self.inputs)
        self.leExcel.returnPressed.connect(self.inputs)
        self.leSheet.returnPressed.connect(self.inputs)
        self.leRows.returnPressed.connect(self.inputs)
        self.leCountry.returnPressed.connect(self.close)
        self.leExcel.returnPressed.connect(self.close)
        self.leSheet.returnPressed.connect(self.close)
        self.leRows.returnPressed.connect(self.close)
        self.btn_Ok.clicked.connect(self.inputs)
        self.btn_Ok.clicked.connect(self.close)
        self.btn_Cancel.clicked.connect(self.on_cancel)
        self.btn_Cancel.clicked.connect(self.close)

        self.show()

    def inputs(self):
        global country_input
        global excel_name_input
        global sheet_num_input
        global rows_num_input

        country_input = self.leCountry.text()
        excel_name_input = f'{self.leExcel.text()}.xlsx'
        sheet_num_input = int(self.leSheet.text())
        rows_num_input = int(self.leRows.text())

    # Create UI
    def setup_UI(self):
        self.setWindowTitle('Signals And Slots')
        self.setGeometry(500, 300, 600, 500)

        self.create_form_groupbox()
        self.create_buttons()

        # Create main Layout
        main_layout = qtw.QVBoxLayout(self)
        main_layout.addWidget(self.form_groupbox)
        main_layout.addLayout(self.buttons_layout)

    # Create Form Group Box
    def create_form_groupbox(self):
        # Create For Group Box
        self.leCountry = qtw.QLineEdit(self)
        self.leExcel = qtw.QLineEdit(self)
        self.leSheet = qtw.QLineEdit(self)
        self.leRows = qtw.QLineEdit(self)

        self.form_groupbox = qtw.QGroupBox('Login Form')
        self.form_layout = qtw.QFormLayout(self)
        self.form_groupbox.setLayout(self.form_layout)

        self.form_layout.addRow('Country code:', self.leCountry)
        self.form_layout.addRow('Excel file name:', self.leExcel)
        self.form_layout.addRow('Sheet (0: SCANNING; 1: AUDIT; 2: A2S):', self.leSheet)
        self.form_layout.addRow('Rows:', self.leRows)

    # Create Buttons Layout
    def create_buttons(self):
        self.buttons_layout = qtw.QHBoxLayout()
        self.btn_Ok = qtw.QPushButton('OK')
        self.btn_Cancel = qtw.QPushButton('Cancel')
        self.buttons_layout.addWidget(self.btn_Ok)
        self.buttons_layout.addWidget(self.btn_Cancel)

    def on_cancel(self):
        print('Canceled!')