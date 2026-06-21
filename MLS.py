


import os
from openpyxl import Workbook
from tkinter import messagebox
from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import datetime, timedelta
#import about window
from about_ui import Ui_About_window
from threading import *



#get interest rate
with open("Interest_set.txt", 'r') as f:
    f2 = f.read()
    for x in f2:
        if x == "[":
            c1 = f2.index(x)
        if x == "]":
            c2 = f2.index(x)
    interrest = f2[c1+1:c2]



#confirm window class
class Ui_Dialog(object):

    def saveToXlsx(self):

        self.Dialog.hide()

        wb = Workbook()
        sheet = wb.active


        for i in [deta]:
            sheet.append(i)
        wb.save("raw.xlsx")

        os.popen("script.vbs")

        messagebox.showinfo("", "Data saved successfully")

    def decline(self):
        self.Dialog.hide()


        

    def setupUi(self, Dialog):
        self.Dialog = Dialog
        Dialog.setObjectName("Dialog")
        Dialog.resize(616, 615)
        Dialog.setMaximumWidth(616)
        Dialog.setMinimumWidth(616)
        Dialog.setMaximumHeight(615)
        Dialog.setMinimumHeight(615)
        Dialog.setWindowIcon(QtGui.QIcon('elements/app.png'))



        Dialog.setStyleSheet("""
            QDialog {
                background-color: #87CEEB;
            }
        """)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(30, 0, 371, 71))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_9 = QtWidgets.QLabel(Dialog)
        self.label_9.setGeometry(QtCore.QRect(30, 360, 121, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.line_2 = QtWidgets.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(20, 510, 571, 16))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.line_2.setFont(font)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(30, 130, 321, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_5 = QtWidgets.QLabel(Dialog)
        self.label_5.setGeometry(QtCore.QRect(30, 190, 521, 31))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.pushButton_2 = QtWidgets.QPushButton(Dialog)
        self.pushButton_2.setGeometry(QtCore.QRect(298, 540, 111, 31))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(170, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.pushButton_2.setPalette(palette)
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setStyleSheet("QPushButton{\n"
"border-radius : 10px;\n"
"background-color : white;\n""}")
        self.pushButton_2.clicked.connect(self.decline)
        
        #self.pushButton_2.clicked.connect()

        self.label_4 = QtWidgets.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(30, 160, 231, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_6 = QtWidgets.QLabel(Dialog)
        self.label_6.setGeometry(QtCore.QRect(30, 220, 391, 31))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(180, 540, 101, 31))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.pushButton.setPalette(palette)
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setStyleSheet("QPushButton{\n"
"border-radius : 10px;\n"
"background-color : white;\n"
"}")
        self.pushButton.clicked.connect(self.saveToXlsx)
        ##
        self.label_7 = QtWidgets.QLabel(Dialog)
        self.label_7.setGeometry(QtCore.QRect(30, 300, 221, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.label_12 = QtWidgets.QLabel(Dialog)
        self.label_12.setGeometry(QtCore.QRect(30, 450, 221, 16))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_12.setFont(font)
        self.label_12.setObjectName("label_12")
        self.label_8 = QtWidgets.QLabel(Dialog)
        self.label_8.setGeometry(QtCore.QRect(30, 330, 241, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.label_13 = QtWidgets.QLabel(Dialog)
        self.label_13.setGeometry(QtCore.QRect(30, 480, 221, 16))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_13.setFont(font)
        self.label_13.setObjectName("label_13")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(30, 90, 351, 31))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_10 = QtWidgets.QLabel(Dialog)
        self.label_10.setGeometry(QtCore.QRect(30, 390, 151, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.line = QtWidgets.QFrame(Dialog)
        self.line.setGeometry(QtCore.QRect(20, 270, 571, 16))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.line.setFont(font)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label_11 = QtWidgets.QLabel(Dialog)
        self.label_11.setGeometry(QtCore.QRect(30, 420, 201, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Confirm"))
        self.label.setText(_translate("Dialog", "Are you sure you want to save \n"
        "the following information?"))
        self.label_2.setText(_translate("Dialog", f"Name: {name}"))
        self.label_3.setText(_translate("Dialog", f"Phone Number: {mob_number}"))
        self.label_4.setText(_translate("Dialog", f"Date of birth: {DOB}"))
        self.label_5.setText(_translate("Dialog", f"Address: {address}"))
        self.label_6.setText(_translate("Dialog", f"Profession: {profession}"))
        self.label_7.setText(_translate("Dialog", f"Principal: {PrincipalAmt}"))
        self.label_8.setText(_translate("Dialog", f"Type of item: {item}"))
        self.label_9.setText(_translate("Dialog", f"Purity: {purity}"))
        self.label_10.setText(_translate("Dialog", f"Weight: {weight}"))
        self.label_11.setText(_translate("Dialog", f"Time period: {time_period}"))
        self.label_12.setText(_translate("Dialog", f"Return time: {Return_period}"))
        self.label_13.setText(_translate("Dialog", f"Total Amount: {AMT}"))
        self.pushButton.setText(_translate("Dialog", "Yes"))
        self.pushButton_2.setText(_translate("Dialog", "No"))


class Ui_MLS(object):

    

    def openAboutWindow(self):
        self.window = QtWidgets.QDialog()
        self.ui = Ui_About_window()
        self.ui.setupUi(self.window)
        self.window.show() 

    def confirm(self):
        
        self.window2 = QtWidgets.QDialog()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self.window2)
        self.window2.show()


        
        

    
    def prepData(self):
        global name, mob_number, DOB, address, profession, PrincipalAmt,item, purity, weight, issued_date, time_period, Return_period, AMT 
        global deta

        name = self.NameBox.text()        
        mob_number = self.phNumber.text()                    
        DOB = self.dateEdit.date().toPyDate()        
        address = self.AddressBox.toPlainText()
        profession = self.Prof_combo.currentText()
        PrincipalAmt = self.PriceBox.text()
        item = self.item_combo.currentText()
        purity = self.purity_spinbox.value()
        weight = self.WeightBox.text()        
        issued_date = datetime.today().date()
        time_period = f"{self.timeBox.text()} {self.time_combo.currentText()}"
        time_period_num = self.timeBox.text()
        
    

        if any(x == "" for x in [name, mob_number, DOB, address, profession, PrincipalAmt, item, purity, weight, time_period]):
            messagebox.showerror("ERROR", "Please fill all spaces.")

        elif not name.isalpha() and not " " in name:
            messagebox.showerror("ERR0R", "Wrong customer name.")

        elif not mob_number.isnumeric():
            messagebox.showerror("ERR0R", "Wrong phone number.")

        elif len(mob_number) != 10:
            messagebox.showerror("ERR0R", "Phone number must be 10 digits.")

            
        elif not PrincipalAmt.isnumeric():
            messagebox.showerror("ERR0R", "Wrong principal")
            
        elif not weight.isnumeric():
            messagebox.showerror("ERR0R", "Enter correct weight")
            
        elif not time_period_num.isnumeric() or int(time_period_num) == 0:
            messagebox.showerror("ERR0R", "Wrong time period")
            
        else:    

           
        
            if self.time_combo.currentText() == "Months":
                Days = int(time_period_num) * 30 
            if self.time_combo.currentText() == "Years":
                Days = int(time_period_num) * 360 
            if self.time_combo.currentText() == "Days":
                Days = int(time_period_num) * 1


            Return_period = (datetime.today() + timedelta(days=Days)).date()

            Interest = ( float(PrincipalAmt) * (float(Days)*0.002738) * float(interrest) ) / 100
            Interest = format(Interest, ".2f")
            AMT = int(PrincipalAmt) + float(Interest) 
          

            deta = (name, int(mob_number), DOB, address, profession, float(PrincipalAmt), int(AMT), item, purity, int(weight), time_period, issued_date, Return_period)
        
            
            self.confirm()

        

    def launch_Record(self):
        os.popen("loanData.xlsm")
        
    def run(self):
        t1  = Thread(target=self.launch_Record)
        t1.start()
        


    def setupUi(self, MLS):
        MLS.setObjectName("MLS")
        MLS.resize(925, 754)
        MLS.setMaximumWidth(925)
        MLS.setMinimumWidth(925)
        MLS.setMaximumHeight(754)
        MLS.setMinimumHeight(754)
        MLS.setWindowIcon(QtGui.QIcon('elements/app.png'))
        MLS.setStyleSheet("""
            QDialog {
                background-color: #abd0f5;
            }
        """)
        self.Title = QtWidgets.QLabel(MLS)
        self.Title.setGeometry(QtCore.QRect(160, 0, 921, 81))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 85, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(131, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.NoRole, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ToolTipText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 85, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(131, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.NoRole, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ToolTipText, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 85, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(131, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(131, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 255, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.NoRole, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ToolTipText, brush)
        self.Title.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Papyrus")
        font.setPointSize(19)
        font.setBold(True)
        font.setWeight(75)
        self.Title.setFont(font)
        self.Title.setText("")
        self.Title.setPixmap(QtGui.QPixmap("elements/title.png"))
        self.Title.setObjectName("Title")
        self.Name_label = QtWidgets.QLabel(MLS)
        self.Name_label.setGeometry(QtCore.QRect(40, 150, 111, 41))
        self.Name_label.setObjectName("Name_label")
        self.line = QtWidgets.QFrame(MLS)
        self.line.setGeometry(QtCore.QRect(-10, 70, 981, 31))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label_5 = QtWidgets.QLabel(MLS)
        self.label_5.setGeometry(QtCore.QRect(430, 90, 75, 16))
        self.label_5.setStyleSheet("""
    QLabel {
        color: #1f4e79;
        font-size: 20px;
        font-weight: bold;
    }
""")
        self.label_5.setObjectName("label_5")
        self.line_2 = QtWidgets.QFrame(MLS)
        self.line_2.setGeometry(QtCore.QRect(0, 110, 971, 16))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.Address_label = QtWidgets.QLabel(MLS)
        self.Address_label.setGeometry(QtCore.QRect(60, 290, 55, 16))
        self.Address_label.setObjectName("Address_label")
        self.Profession_label = QtWidgets.QLabel(MLS)
        self.Profession_label.setGeometry(QtCore.QRect(60, 340, 81, 16))
        self.Profession_label.setObjectName("Profession_label")
        self.Prof_combo = QtWidgets.QComboBox(MLS)
        self.Prof_combo.setGeometry(QtCore.QRect(140, 330, 181, 31))
        self.Prof_combo.setObjectName("Prof_combo")
        self.Prof_combo.addItem("")
        self.Prof_combo.setItemText(0, "")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.addItem("")
        self.Prof_combo.setItemText(8, "")
        self.Prof_combo.setStyleSheet("QListView""{"" background-color: orange ""}" )
        self.line_3 = QtWidgets.QFrame(MLS)
        self.line_3.setGeometry(QtCore.QRect(-10, 410, 971, 16))
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.line_4 = QtWidgets.QFrame(MLS)
        self.line_4.setGeometry(QtCore.QRect(-20, 370, 981, 31))
        self.line_4.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.label_7 = QtWidgets.QLabel(MLS)
        self.label_7.setGeometry(QtCore.QRect(410, 390, 161, 21))
        self.label_7.setStyleSheet("""
    QLabel {
        color: #1f4e79;
        font-size: 20px;
        font-weight: bold;
    }
""")
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(MLS)
        self.label_8.setGeometry(QtCore.QRect(50, 550, 121, 16))
        self.label_8.setObjectName("label_8")
        self.purity_spinbox = QtWidgets.QSpinBox(MLS)
        self.purity_spinbox.setGeometry(QtCore.QRect(170, 541, 71, 31))
        self.purity_spinbox.setMinimum(15)
        self.purity_spinbox.setMaximum(24)
        self.purity_spinbox.setObjectName("purity_spinbox")
        self.label_9 = QtWidgets.QLabel(MLS)
        self.label_9.setGeometry(QtCore.QRect(50, 580, 111, 16))
        self.label_9.setObjectName("label_9")
        self.WeightBox = QtWidgets.QLineEdit(MLS)
        self.WeightBox.setGeometry(QtCore.QRect(170, 580, 91, 22))
        self.WeightBox.setObjectName("WeightBox")
        self.label_10 = QtWidgets.QLabel(MLS)
        self.label_10.setGeometry(QtCore.QRect(50, 499, 81, 16))
        self.label_10.setObjectName("label_10")
        self.item_combo = QtWidgets.QComboBox(MLS)
        self.item_combo.setGeometry(QtCore.QRect(170, 500, 141, 31))
        self.item_combo.setObjectName("item_combo")
        self.item_combo.addItem("")
        self.item_combo.setItemText(0, "")
        self.item_combo.addItem("")
        self.item_combo.addItem("")
        self.item_combo.addItem("")
        self.item_combo.addItem("")
        self.item_combo.addItem("")
        self.item_combo.setStyleSheet("QListView""{"" background-color: orange ""}" )
        self.label_11 = QtWidgets.QLabel(MLS)
        self.label_11.setGeometry(QtCore.QRect(50, 620, 101, 16))
        self.label_11.setObjectName("label_11")
        self.DOB_label = QtWidgets.QLabel(MLS)
        self.DOB_label.setGeometry(QtCore.QRect(60, 250, 91, 16))
        self.DOB_label.setObjectName("DOB_label")
        self.dateEdit = QtWidgets.QDateEdit(MLS)
        self.dateEdit.setGeometry(QtCore.QRect(140, 250, 110, 22))
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setObjectName("dateEdit")
        self.timeBox = QtWidgets.QLineEdit(MLS)
        self.timeBox.setGeometry(QtCore.QRect(170, 620, 81, 22))
        self.timeBox.setObjectName("timeBox")
        self.time_combo = QtWidgets.QComboBox(MLS)
        self.time_combo.setGeometry(QtCore.QRect(260, 620, 73, 22))
        self.time_combo.setObjectName("time_combo")
        self.time_combo.addItem("")
        self.time_combo.addItem("")
        self.time_combo.addItem("")
        self.label_2 = QtWidgets.QLabel(MLS)
        self.label_2.setGeometry(QtCore.QRect(50, 440, 131, 20))
        self.label_2.setObjectName("label_2")
        self.PriceBox = QtWidgets.QLineEdit(MLS)
        self.PriceBox.setGeometry(QtCore.QRect(190, 440, 113, 22))
        self.PriceBox.setObjectName("PriceBox")
        self.label_14 = QtWidgets.QLabel(MLS)
        self.label_14.setGeometry(QtCore.QRect(50, 480, 211, 16))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        self.label_14.setPalette(palette)
        self.About = QtWidgets.QPushButton(MLS)
        self.About.setGeometry(QtCore.QRect(0, 0, 61, 31))
        self.About.setObjectName("About")
        self.About.clicked.connect(self.openAboutWindow)

        self.NameBox = QtWidgets.QLineEdit(MLS)
        self.NameBox.setGeometry(QtCore.QRect(140, 160, 291, 31))
        self.NameBox.setObjectName("NameBox")
        self.AddressBox = QtWidgets.QTextEdit(MLS)
        self.AddressBox.setGeometry(QtCore.QRect(140, 280, 271, 31))
        self.AddressBox.setObjectName("AddressBox")
        self.savebutton = QtWidgets.QPushButton(MLS)
        self.savebutton.setGeometry(QtCore.QRect(310, 690, 111, 41))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(30, 45, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(93, 29, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(30, 45, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(93, 29, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.savebutton.setPalette(palette)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.savebutton.setFont(font)
        self.savebutton.setObjectName("savebutton")
        self.savebutton.setStyleSheet("QPushButton{\n"
"                    background-color:rgb(0, 0, 255);\n"
"                    border-radius:15px;\n"
"                    color:white;\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"background-color:rgb(255, 255, 255);\n"
"border: 3px solid blue;\n"
"color:blue;\n"
"\n"
"}")
        self.savebutton.clicked.connect(self.prepData)
        self.line_5 = QtWidgets.QFrame(MLS)
        self.line_5.setGeometry(QtCore.QRect(10, 650, 891, 20))
        self.line_5.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_5.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_5.setObjectName("line_5")
        self.label = QtWidgets.QLabel(MLS)
        self.label.setGeometry(QtCore.QRect(60, 210, 91, 16))
        self.label.setObjectName("label")
        self.phNumber = QtWidgets.QLineEdit(MLS)
        self.phNumber.setGeometry(QtCore.QRect(160, 210, 151, 22))
        self.phNumber.setObjectName("phNumber")
        self.openRecord = QtWidgets.QPushButton(MLS)
        self.openRecord.setGeometry(QtCore.QRect(490, 690, 121, 41))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(85, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(85, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.openRecord.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Dubai")
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.openRecord.setFont(font)
        self.openRecord.setObjectName("openRecord")
        self.openRecord.setStyleSheet("QPushButton{\n"
"border-radius : 15px;\n"
"color: white;\n"
"background-color: rgb(34, 225, 88)\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"border-radius : 15px;\n"
"border : 3px solid rgb(34, 225, 88);\n"
"color : rgb(34, 225, 88);\n"
"background-color: white;\n"
"}")
        self.openRecord.clicked.connect(self.run)

        self.retranslateUi(MLS)
        QtCore.QMetaObject.connectSlotsByName(MLS)

    def retranslateUi(self, MLS):
        _translate = QtCore.QCoreApplication.translate
        MLS.setWindowTitle(_translate("MLS", "MLS V1.0"))
        self.Name_label.setText(_translate("MLS", "     Full Name"))
        self.label_5.setText(_translate("MLS", "Profile"))
        self.Address_label.setText(_translate("MLS", "Address"))
        self.Profession_label.setText(_translate("MLS", "Profession"))
        self.Prof_combo.setItemText(1, _translate("MLS", "Farmer"))
        self.Prof_combo.setItemText(2, _translate("MLS", "Business"))
        self.Prof_combo.setItemText(3, _translate("MLS", "Job"))
        self.Prof_combo.setItemText(4, _translate("MLS", "Docter"))
        self.Prof_combo.setItemText(5, _translate("MLS", "Lawyer"))
        self.Prof_combo.setItemText(6, _translate("MLS", "Engineer/Architect"))
        self.Prof_combo.setItemText(7, _translate("MLS", "Shopkeeper"))
        self.label_7.setText(_translate("MLS", "Morgage details"))
        self.label_8.setText(_translate("MLS", "Purity(in karrats)"))
        self.label_9.setText(_translate("MLS", "Weight(in grams)"))
        self.label_10.setText(_translate("MLS", "Type of item"))
        self.item_combo.setItemText(1, _translate("MLS", "Ring"))
        self.item_combo.setItemText(2, _translate("MLS", "Necklace"))
        self.item_combo.setItemText(3, _translate("MLS", "Earrings"))
        self.item_combo.setItemText(4, _translate("MLS", "Bangals"))
        self.item_combo.setItemText(5, _translate("MLS", "Other"))
        self.label_11.setText(_translate("MLS", "Time period"))
        self.DOB_label.setText(_translate("MLS", "Date of birth"))
        self.time_combo.setItemText(0, _translate("MLS", "Days"))
        self.time_combo.setItemText(1, _translate("MLS", "Months"))
        self.time_combo.setItemText(2, _translate("MLS", "Years"))
        self.label_2.setText(_translate("MLS", "Principal Amount(IND)"))
        self.label_14.setText(_translate("MLS", f"Note:- We apply {interrest}% interest"))
        self.About.setText(_translate("MLS", "About"))
        self.savebutton.setText(_translate("MLS", "Save"))
        self.label.setText(_translate("MLS", "Mobile number"))
        self.openRecord.setText(_translate("MLS", "Open record"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MLS = QtWidgets.QDialog()
    ui = Ui_MLS()
    ui.setupUi(MLS)
    MLS.show()
    sys.exit(app.exec_())
