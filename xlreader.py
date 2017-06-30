import openpyxl

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QFileDialog

# Usage instructions for dumbos - usually not needed
#print('Welcome to xlreader, created by Benjamin Steenhoek.')
#print('If you have used this utility before, please note that any generated')
#print('workbooks should be CLOSED before running this program to prevent errors.\n')
#wbname = input('Please input workbook name, then press enter. ')

# Start program
#try:
wb = openpyxl.load_workbook(wbname)
#except FileNotFoundError:
#    print('File "' + wbname + '" was not found. Check your spelling and capitalization and try again.')
#    input('Press "Enter" to close the program.')

asins = {}

# Open workbook
ws1 = wb.get_sheet_by_name('Sheet1')
for row in range(2,ws1.max_row):
    asin_cell = "C" + str(row)
    asin = ws1[asin_cell].value
    quantity_cell = "E" + str(row)
    quantity = ws1[quantity_cell].value
    if asin not in asins.keys():  # New value
        asins[asin] = list()
        asins[asin].append(ws1["A" + str(row)].value)
        asins[asin].append(ws1["B" + str(row)].value)
        asins[asin].append(ws1["D" + str(row)].value)
        asins[asin].append(quantity)
        asins[asin].append(ws1["F" + str(row)].value)
        asins[asin].append(ws1["G" + str(row)].value)
    else:
        asins[asin][3] += quantity   # Add to old value

# Save to new workbook
newwb = openpyxl.Workbook()
ws = newwb.active
ws.title = 'CorrectedASINS'

ws["A1"] = "FC"
ws["B1"] = "Category"
ws["C1"] = "ASIN"
ws["D1"] = "Description"
ws["E1"] = "Units"
ws["F1"] = "Cost"
ws["G1"] = "Ext Cost"

count = 2
for asin in asins.keys():
    ws["A" + str(count)] = asins[asin][0]
    ws["B" + str(count)] = asins[asin][1]
    ws["C" + str(count)] = asin
    ws["D" + str(count)] = asins[asin][2]
    ws["E" + str(count)] = asins[asin][3]
    ws["F" + str(count)] = asins[asin][4]
    ws["G" + str(count)] = asins[asin][5]
    count += 1

newwbname = wbname[:-5] + '_orderedbyASINs' + ".xlsx"
try:
    newwb.save(newwbname)
except PermissionError:
    print('Please close ' + newwbname + ', then try again.')
input('Workbook saved in xlreader folder to name ' + newwbname + '. Press "Enter" to close.')

# GUI
class LoginWidget(QWidget):

    StartTime = None
    EndTime = None

    Username = ""
    Liquidations = 0

    def __init__(self):
        super().__init__()

        self.initialize_gui()
        self.connect_login_db()

    def initialize_gui(self):
        """
        Initialize the window.
        """
        self.setWindowTitle('Roney Industries Employee Login')
        self.setWindowIcon(QIcon("not_connected.png"))

        # Username/Password Input
        Label = QLabel(self)
        Label.setText("Enter username:")

        self.XLInput = QLineEdit(self)
        self.SubmitXL = QLineEdit(self)

        self.fileName = QFileDialog.getOpenFileName(self,"Open Excel", "C:")

    def openXL(self):
        xlname = self.XLInput.text

