# -*- coding: utf-8 -*-
# 2022/08/23 初版
# 2022/08/24 正式發行v1.0 新增使用者模糊查詢
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QWidget,
                            QGridLayout,  QMessageBox,
                             QTableWidget,QApplication,QHBoxLayout)
import qdarkstyle, sys, os
import pyodbc
class Ui_Dialog(QMainWindow):
    def __init__(self):
        super().__init__()

        #給予查詢單號和使用人的暫存清單
        self.direct = []
        self.direct_user = []
        #給予暫存帳號密碼的暫存清單
        self.info = []
        self.info2= []
        self.connection = []
        self.cursor = []

        self.setObjectName("Dialog")
        self.resize(1140, 850)
        #加狀態列 顯示現在使用者是誰
        self.status = self.statusBar()
        self.status.showMessage(os.environ.get("Username")+"  Hi Bro")

        # 視窗標題
        self.setWindowTitle("資料庫查詢")
        # 資料庫連線connect變數
        self.xo =""
        self.setWindowOpacity(0.9)
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(220, 10, 311, 91))
        self.label.setObjectName("label")
        # 賦予表格物件
        self.tableWidget = QtWidgets.QTableWidget(self)
        self.tableWidget.setGeometry(QtCore.QRect(270, 121, 831, 311))
        self.tableWidget.setObjectName("tableWidget")

        self.left_widget = QWidget()
        self.left_widget.setObjectName('left_widget')
        self.left_layout = QGridLayout()
        self.left_widget.setLayout(self.left_layout)

        self.tableWidget_2 = QtWidgets.QTableWidget(self)
        self.tableWidget_2.setGeometry(QtCore.QRect(270, 471, 831, 311))
        self.tableWidget_2.setObjectName("tableWidget_2")
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(0, item)
        item.setText("Name")
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(3, item)

        # 賦予單號輸入欄位和使用人輸入欄位物件
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.lineEdit.setPlaceholderText("輸入單號模式查詢(Enter)")
        self.lineEdit.setToolTip("輸入單號後 按Enter會跳出一個提示\n輸入範例：20220000000")
        self.lineEdit.returnPressed.connect(self.kkey)
        self.lineEdit.setGeometry(QtCore.QRect(40, 89, 171, 31))
        self.lineEdit.setObjectName("lineEdit")
        self.lineuser = QtWidgets.QLineEdit(self)
        self.lineuser.setPlaceholderText("輸入使用者名稱模式查詢(Enter)")
        self.lineuser.setToolTip("輸入使用者名稱後 按Enter會跳出一個提示\n輸入範例：dathan is handsome")
        self.lineuser.returnPressed.connect(self.kkuser)
        self.lineuser.setGeometry(QtCore.QRect(40, 130, 171, 31))
        self.lineuser.setObjectName("lineEdituser")
        # 賦予左側按鈕-查詢單號-查詢使用者

        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(40, 180, 171, 41))
        self.pushButton.setText("單號查詢")
        self.pushButton.setToolTip("輸入單號後 按此按鈕會顯示在右側表格內")

        self.pushButton.clicked.connect(self.loaddata)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self)
        self.pushButton_2.setGeometry(QtCore.QRect(40, 230, 171, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setText("使用者查詢")
        self.pushButton_2.setToolTip("輸入使用者名稱後 按此按鈕後會顯示在右側表格內")

        self.pushButton_2.clicked.connect(self.loaddata_User)
        self.pushButton_3 = QtWidgets.QPushButton(self)
        self.pushButton_3.setGeometry(QtCore.QRect(40, 280, 171, 41))
        self.pushButton_3.setObjectName("pushButton_3")
        self.buttonLayout = QHBoxLayout(self.left_widget)
        self.buttonLayout.addWidget(self.pushButton_3)
        self.pushButton_3.clicked.connect(self.gone)
        self.pushButton_3.setText("重設表格")

        # 定義視窗介面黑色風格
        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
        # tablewidgt to view data
        self.query_result = QTableWidget()
        self.left_layout.addWidget(self.query_result, 9, 0, 2, 5)
        self.label.setText("資料庫查詢V1.0")
        self.label.setFont(QtGui.QFont("Times", 24, QtGui.QFont.Bold))

        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setGeometry(QtCore.QRect(20, 470, 261, 41))
        self.label_2.setStyleSheet(
            "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:0, y2:0.715909, stop:0 rgba(0, 0, 0, 9), stop:0.375 rgba(0, 0, 0, 50), stop:0.835227 rgba(0, 0, 0, 75));\n"
            "border-radius:20px;")
        self.label_2.setText("")
        self.label_2.setObjectName("label_2")
        self.label_4 = QtWidgets.QLabel(self)
        self.label_4.setGeometry(QtCore.QRect(60, 490, 190, 140))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(90)

        self.label_4.setFont(font)
        self.label_4.setStyleSheet("color:rgba(255, 255, 255, 210);")
        self.label_4.setObjectName("label_4")
        # 賦予登入輸入框
        self.lineEdit_2 = QtWidgets.QLineEdit(self)
        self.lineEdit_2.setGeometry(QtCore.QRect(30, 635, 200, 40))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setStyleSheet("background-color:rgba(0, 0, 0, 0);\n"
                                      "border:none;\n"
                                      "border-bottom:2px solid rgba(105, 118, 132, 255);\n"
                                      "color:rgba(255, 255, 255, 230);\n"
                                      "padding-bottom:7px;")
        self.lineEdit_2.setEchoMode(QtWidgets.QLineEdit.Password)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self)
        self.lineEdit_3.setGeometry(QtCore.QRect(30, 570, 200, 40))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setText(("Log In"))
        self.lineEdit_2.setPlaceholderText("  Password")
        self.lineEdit_3.setPlaceholderText("  User Name")
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setStyleSheet("background-color:rgba(0, 0, 0, 0);\n"
                                      "border:none;\n"
                                      "border-bottom:2px solid rgba(105, 118, 132, 255);\n"
                                      "color:rgba(255, 255, 255, 230);\n"
                                      "padding-bottom:7px;")
        self.lineEdit_3.setObjectName("lineEdit_3")
        #登入按鈕
        self.pushButton_3 = QtWidgets.QPushButton(self)
        self.pushButton_3.setText("L o g  I n")
        self.pushButton_3.clicked.connect(self.loginSql)
        self.pushButton_3.setGeometry(QtCore.QRect(30, 710, 200, 40))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        #輸錯密碼提示
        self.error = QtWidgets.QLabel(self)
        self.error.setGeometry(QtCore.QRect(30, 690, 201, 16))
        self.error.setText("")
        self.error.setStyleSheet("color:rgba(255, 85, 0);")
        self.error.setObjectName("error")
        QtCore.QMetaObject.connectSlotsByName(self)

    def loginSql(self,con):
        self.error.setText("")
        ps_text = self.lineEdit_3.text()
        ac_text = self.lineEdit_2.text()
        if len(ac_text)==0 or len(ps_text)==0:
            self.error.setText("Please input all fields.")

        try:
            global xo, conn, cu, st
            con = self.connection
            cu = self.cursor
            con.append(pyodbc.connect('Driver=SQL Server;' 'Server=10.X.X.X;'
                                      'Database=OA;'  'uid={};' 'pwd={};'.format(ps_text, ac_text)).cursor())#根據您的實際資訊去修改IP及資料庫名稱
            xo = pyodbc.connect('Driver=SQL Server;' 'Server=10.X.X.X;'
                                'Database=OA;'  'uid={};' 'pwd={};'.format(ps_text, ac_text))#根據您的實際資訊去修改IP及資料庫名稱
            cu.append("con.cursor()")
            conn = con
        except:
            self.error.setText("Please input all fields or correct password")
        if con:
            print('connect to SQL Server successfully')

            self.lineEdit_3.setEnabled(False)
            self.lineEdit_2.setEnabled(False)
            self.lineEdit_2.hide()
            self.lineEdit_3.hide()
            self.pushButton_3.hide()
            self.label_4.hide()
            self.label_4.show()
            self.label_4.move(40, 490)
            self.label_4.setText("connect to SQL Server \nsuccessfully")
            return con
        else:
            print('connection failed')
            self.label_4.hide()
            self.label_4.show()
            self.label_4.setText("L o g i n")
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_2.show()
            self.lineEdit_3.show()
            self.pushButton_3.show()
            return False


    #暫時未使用
    def cklogin(self):
        import subprocess

        result = subprocess.getoutput('netstat -an | findstr 1433')
        print(result)

        if result:
            print('connect to SQL Server successfully')
            self.lineEdit_3.setEnabled(False)
            self.lineEdit_2.setEnabled(False)
            self.lineEdit_2.hide()
            self.lineEdit_3.hide()
            self.pushButton_3.hide()
            return True
        else:
            print('connection failed')
            self.lineEdit_2.show()
            self.lineEdit_3.show()

            return False
    #重設表格
    def gone(self):

        self.tableWidget.clearContents()

        self.tableWidget_2.clearContents()
    #單號輸入框
    def kkey(self):
        try:

            line_appnum = self.lineEdit.text()
            if len(line_appnum) == 0 or len(line_appnum) <= 10:
                print('輸入結果：未符合長度')
                QMessageBox.about(self, "輸入結果", "<font size='26' color='red'>未符合長度，請確認單號是否輸入正確</font>")
            else:


                self.direct=(line_appnum)
                QMessageBox.information(self, "已輸入單號：", "<font size='26' color='green'>{}</font>".format(self.direct + "\n"))
                self.lineEdit.clear()
                self.direct_user.clear()
                if len(self.direct_user) == 0:
                    self.lineuser.setEnabled(False)
                    self.lineuser.hide()


        except:
            pass
    #使用人輸入框
    def kkuser(self):
        try:

            line2_appnum = self.lineuser.text()
            if len(line2_appnum) == 0 or len(line2_appnum) <= 1:
                print('輸入結果：未符合長度')
                QMessageBox.about(self, "輸入結果", "<font size='26' color='red'>未符合長度，請確認使用者是否輸入正確</font>")
            else:

                self.direct_user = (line2_appnum)
                QMessageBox.information(self, "已輸入使用者名稱：",
                                        "<font size='26' color='green'>{}</font>".format(self.direct_user + "\n"))
                self.lineuser.clear()
                self.direct.clear()
                if len(self.direct) == 0:
                    self.lineEdit.setEnabled(False)
                    self.lineEdit.hide()

        except:
            pass

    #查詢單號表資料庫事件
    def loaddata(self):
        global sqlconA,sqlconB,querycolumA,querycolumB
        try:
            self.setWindowTitle("資料庫查詢   --- 正在查詢{}".format(self.direct))
            sqlconA = xo.cursor()

            querycolumA = sqlconA.execute("SELECT AppNum,Applicant FROM NDA Where AppNum = '{}'".format(self.direct))#根據您的資料庫資料表欄位名稱
            queryrowsA = querycolumA.fetchone()
            sqlconA.execute("SELECT AppNum,Applicant FROM NDA Where AppNum = '-{}'".format(self.direct))#根據您的資料庫資料表欄位名稱
            queryresultA = querycolumA.fetchall()
            self.tableWidget.setColumnCount(len(queryrowsA))
            self.tableWidget.setRowCount(len(queryresultA))


            sqlconB = xo.cursor()

            querycolumB = sqlconB.execute("SELECT AppNum,Enabled,UserName,RemoteType,StartDate,EndDate,AddFlag,"
                                          "DelFlag,GroupChk FROM NDB Where AppNum = '{}'".format(self.direct))#根據您的資料庫資料表欄位名稱
            queryrowsB = querycolumB.fetchone()
            sqlconB.execute("SELECT AppNum,Enabled,UserName,RemoteType,StartDate,EndDate,AddFlag,"
                                          "DelFlag,GroupChk FROM NDB Where AppNum = '{}'".format(self.direct))#根據您的資料庫資料表欄位名稱
            queryresultB = querycolumB.fetchall()
            self.tableWidget_2.setColumnCount(len(queryrowsB))
            self.tableWidget_2.setRowCount(len(queryresultB))
            ls = ["AppNum","Applicant"]#根據您的資料庫資料表欄位名稱
            ls2 = ["AppNum","Enabled","UserName","RemoteType","StartDate","EndDate","AddFlag","DelFlag","GroupChk"]#根據您的資料庫資料表欄位名稱
            self.tableWidget.setHorizontalHeaderLabels(ls)
            self.tableWidget_2.setHorizontalHeaderLabels(ls2)
            row = 0
            colcount = 0
            row2 = 0
            colcount2 = 0
            for person in sqlconA.execute("SELECT AppNum,Applicant FROM NDA Where AppNum = '{}'".format(self.direct)):#根據您的資料庫資料表欄位名稱
                personlist = []

                for count in person:
                    personlist.append(str(count))

                    self.tableWidget.setItem(row, colcount, QtWidgets.QTableWidgetItem(count))


                    self.tableWidget.setItem(row, colcount, QtWidgets.QTableWidgetItem(str(count)))
                    colcount += 1
            for person2 in sqlconB.execute("SELECT AppNum,Enabled,UserName,RemoteType,StartDate,EndDate,AddFlag,"
                                          "DelFlag,GroupChk FROM NDB Where AppNum = '{}'".format(self.direct)):#根據您的資料庫資料表欄位名稱
                personlist2 = []

                for count2 in person2:
                    personlist2.append(str(count2))
                    self.tableWidget_2.setItem(row2, colcount2, QtWidgets.QTableWidgetItem(count2))
                    self.tableWidget_2.setItem(row2, colcount2, QtWidgets.QTableWidgetItem(str(count2)))
                    colcount2 += 1
            self.direct = []
            self.lineuser.setEnabled(True)
            self.lineuser.show()
            querycolumA.close()
            querycolumB.close()

        except pyodbc.ProgrammingError as e:
            self.lineEdit_2.show()
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_3.show()
            self.pushButton_3.show()
            self.label_4.hide()
            self.label_4.show()
            self.label_4.setText("Log i n")
            print(e)
            QMessageBox.information(self, "查詢結果：",
                                    "<font size='26' color='red'>{}</font>".format("已中斷查詢，請重新登入" + "\n"))
        except Exception as e2:
            print('其他錯誤或是未登入')
            QMessageBox.information(self, "查詢結果：",
                                    "<p><font size='26' color='red'>{}</font>\n</p>"
                                    "<br><font size='26' color='red'>{}</font></br>"
                                    "<br><font size='26' color='red'>{}</font></br>".format("1. 請確認是否有輸入到單號" ,"2. 單號是否存在？","3. 是否有登入？"))
            print(e2)
            self.direct = []
            self.lineuser.setEnabled(True)
            self.lineuser.show()
    #查詢使用人資料庫事件
    def loaddata_User(self):
        global sqlconA,sqlconB,querycolumA,querycolumB
        try:
            self.setWindowTitle(" 資料庫查詢   --- 正在查詢{}".format(self.direct_user))
            sqlconA = xo.cursor()

            querycolumA = sqlconA.execute("SELECT AppNum,Enabled,UserName,StartDate,EndDate,AddFlag,DelFlag,GroupChk"
                                          " FROM NDB Where CHARINDEX('{}', UserName) >0".format(self.direct_user))#根據您的資料庫資料表欄位名稱
            queryrowsA = querycolumA.fetchone()
            sqlconA.execute("SELECT AppNum,Enabled,UserName,StartDate,EndDate,AddFlag,DelFlag,GroupChk"
                                          " FROM NDB Where CHARINDEX('{}', UserName) >0".format(self.direct_user))#根據您的資料庫資料表欄位名稱
            queryresultA = querycolumA.fetchall()
            self.tableWidget.setColumnCount(len(queryrowsA))
            self.tableWidget.setRowCount(len(queryresultA))


            sqlconB = xo.cursor()

            querycolumB = sqlconB.execute("SELECT UserName,GroupName"
                                          " FROM NDC Where CHARINDEX('{}', UserName) >0".format(self.direct_user))#根據您的資料庫資料表欄位名稱
            queryrowsB = querycolumB.fetchone()
            sqlconB.execute("SELECT UserName,GroupName"
                                          " FROM NDC Where CHARINDEX('{}', UserName) >0".format(self.direct_user))#根據您的資料庫資料表欄位名稱
            queryresultB = querycolumB.fetchall()
            self.tableWidget_2.setColumnCount(len(queryrowsB))
            self.tableWidget_2.setRowCount(len(queryresultB))

            self.tableWidget.setHorizontalHeaderLabels(["AppNum","Enabled","UserName","StartDate","EndDate","AddFlag",
                                                        "DelFlag","GroupChk"])#根據您的資料庫資料表欄位名稱
            self.tableWidget_2.setHorizontalHeaderLabels(["UserName","GroupName"])#根據您的資料庫資料表欄位名稱
            row = 0
            colcount = 0
            row2 = 0
            colcount2 = 0
            for person in sqlconA.execute("SELECT AppNum,Enabled,UserName,StartDate,EndDate,AddFlag,DelFlag,GroupChk"
                                          " FROM NDB Where CHARINDEX('{}', UserName) >0".format(self.direct_user)):#根據您的資料庫資料表欄位名稱
                personlist = []

                for count in person:
                    personlist.append(str(count))

                    self.tableWidget.setItem(row, colcount, QtWidgets.QTableWidgetItem(count))


                    self.tableWidget.setItem(row, colcount, QtWidgets.QTableWidgetItem(str(count)))
                    colcount += 1
            for person2 in sqlconB.execute("SELECT UserName,GroupName"
                                          " FROM NDC Where CHARINDEX('{}', UserName) >0".format(self.direct_user)):#根據您的資料庫資料表欄位名稱
                personlist2 = []

                for count2 in person2:
                    personlist2.append(str(count2))
                    self.tableWidget_2.setItem(row2, colcount2, QtWidgets.QTableWidgetItem(count2))
                    self.tableWidget_2.setItem(row2, colcount2, QtWidgets.QTableWidgetItem(str(count2)))
                    colcount2 += 1
            self.direct_user = []
            self.lineEdit.setEnabled(True)
            self.lineEdit.show()
            querycolumA.close()
            querycolumB.close()

        except pyodbc.ProgrammingError as e:
            self.lineEdit_2.show()
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_3.show()
            self.pushButton_3.show()
            self.label_4.hide()
            self.label_4.show()
            self.label_4.setText("Log i n")
            print(e)
            QMessageBox.information(self, "查詢結果：",
                                    "<font size='26' color='red'>{}</font>".format("已中斷查詢，請重新登入" + "\n"))
        except Exception as e2:
            print('其他錯誤或是未登入')
            QMessageBox.information(self, "查詢結果：",
                                    "<p><font size='26' color='red'>{}</font>\n</p>"
                                    "<br><font size='26' color='red'>{}</font></br>"
                                    "<br><font size='26' color='red'>{}</font></br>".format("1. 請確認是否有輸入到單號" ,"2. 單號是否存在？","3. 是否有登入？"))
            print(e2)
            self.direct_user = []
            self.direct_user = []
            self.lineEdit.setEnabled(True)
            self.lineEdit.show()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    gui = Ui_Dialog()
    #檢測使用者是否符合指定帳號範圍才可使用
    user = os.environ.get("Username")
    users = ['user1','user2']#這邊假設只允許給特定的使用者使用該程式
    if user in users:
        print('確認帳號在准許名單內')
        gui.show()
    else:
        print('該帳號不在准許名單內')
        QMessageBox.warning(None, "查詢結果：", "<p><font size='26' color='red'>本程式只有准許名單內可以開啟</font>\n</p>"
                                           "<br><font size='26' color='red'>#訊息2</font></br>")
        exit()
    sys.exit(app.exec_())