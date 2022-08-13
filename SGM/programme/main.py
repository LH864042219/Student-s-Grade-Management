import os
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QInputDialog
from PyQt5.QtWidgets import QMessageBox
from Mainwindow import Ui_MainWindow
import Add
import pandas as pd
import sqlite3 as sql
import numpy as np


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        # 绘图部分计数及图表标题
        self.count = 0
        self.label = [self.zongmin, self.yuwen, self.shuxue, self.yingyu, self.xuanke1, self.xuanke2, self.xuanke3]
        # 文件地址获取:
        import os
        self.osaddress = os.getcwd()
        self.osaddress = self.osaddress[:26]
        self.list = os.listdir(self.osaddress + '/AllClass')
        # 工具栏添加按钮
        self.toolBar.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        self.toolBar.addAction("添加班级")
        self.addgrade = QtWidgets.QAction("导入新成绩")
        self.delclass = QtWidgets.QAction("删除班级")
        self.delgrade = QtWidgets.QAction("删除成绩")
        self.shuaxin = QtWidgets.QAction("刷新")
        self.huitu = QtWidgets.QAction("绘图")
        self.help = QtWidgets.QAction("帮助")
        self.toolBar.addActions([self.addgrade, self.delgrade, self.delclass, self.shuaxin, self.huitu, self.help])
        self.combobox = QtWidgets.QComboBox()
        self.combobox.addItems(self.list)
        self.toolBar.addWidget(self.combobox)
        self.combobox.currentIndexChanged.connect(self.showinfo)
        # 当工具栏的按钮被点击
        self.toolBar.actionTriggered[QtWidgets.QAction].connect(self.tools)
        # 初始化TableWidgets控件

        self.column = ['姓名    ', '考试场次  ', '总分  ', '总名  ', '语文  ', '数学  ', '英语  ', '理赋  ', '化赋  ', '生赋  ', '政赋  ',
                       '史赋  ', '地赋  ', '技术赋', '语名  ', '数名  ',
                       '英名  ', '理名  ', '化名  ', '生名  ', '政名  ', '史名  ', '地名  ', '技术名  ']
        self.tableWidget.setColumnCount(len(self.column))
        self.tableWidget.setHorizontalHeaderLabels(self.column)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setAlternatingRowColors(True)

    def tools(self, m):  # 工具栏按钮被点击时
        choose = m.text()
        # 当选择添加时
        if choose == '添加班级' or choose == '导入新成绩':
            self.add = Add.Ui_Form()
            self.add.show()
            if choose == '添加班级':
                QMessageBox.information(None, '注意', '导入学生信息时请选择正确的文件,名称处填写学生班级', QMessageBox.Ok)
                self.add.ok.clicked.connect(self.AddClass)
            else:
                QMessageBox.information(None, '注意', '导入成绩时请选择正确的文件,名称处填写本次考试名称', QMessageBox.Ok)
                self.add.ok.clicked.connect(self.AddGrade)
            self.add.pushButton_3.clicked.connect(self.bindList)
        if choose == '删除班级':
            from PyQt5.QtWidgets import QInputDialog
            dbaddress, flag = QInputDialog.getItem(self, "班级", "请选择要删除班级：", self.list, 0, False)
            if flag:
                select = QMessageBox.warning(None, '警告', '您确定要删除' + str(dbaddress) + '吗？',
                                             QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
                if select == QMessageBox.Ok:
                    os.remove(self.osaddress + '/AllClass/' + dbaddress)
                    QMessageBox.information(None, '消息', '删除成功')
        if choose == '删除成绩':
            from PyQt5.QtWidgets import QInputDialog
            dbaddress, flag = QInputDialog.getItem(self, "班级", "请选择班级：", self.list, 0, False)
            if flag:
                db = sql.connect(self.osaddress + '/AllClass/' + dbaddress)
                conn = db.cursor()
                # 读取表名
                conn.execute("select name from sqlite_master where type='table' order by name")
                db.commit()
                stu = conn.fetchall()
                conn.execute("select filename from " + str(stu[0][0]))
                nameall = conn.fetchall()
                allname = []
                for i in range(len(nameall)):
                    allname.append(nameall[i][0])
                testname, flag1 = QInputDialog.getItem(self, "考试", "请选择考试名：", allname, 0, False)
                print(testname)
                if flag1:
                    select = QMessageBox.warning(None, '警告', '您确定要删除' + str(testname) + '吗？',
                                                 QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
                    if select == QMessageBox.Ok:
                        for i in range(len(stu)):
                            conn.execute("delete from " + str(stu[i][0]) + " where filename=?", (testname,))
                            db.commit()
                        conn.close()
                        db.close()
                        QMessageBox.information(None, '消息', '删除成功')
        if choose == '刷新':
            # QMessageBox.information(None, '提示', '功能尚未完成')
            self.list = os.listdir(self.osaddress + '/AllClass')
            self.combobox.addItems(self.list)
            QMessageBox.information(None, '提示', '刷新完成')
        if choose == '绘图':
            from PyQt5.QtWidgets import QInputDialog
            from PyQt5.QtGui import QPixmap
            import matplotlib.pyplot as plt
            # 显示中文
            import matplotlib
            import shutil
            matplotlib.rcParams['font.sans-serif'] = ['SimHei']
            matplotlib.rcParams['axes.unicode_minus'] = False

            dbaddress = self.combobox.currentText()
            db = sql.connect(self.osaddress + '/AllClass/' + dbaddress)
            conn = db.cursor()
            # 读取表内学生名
            conn.execute("select name from sqlite_master where type='table' order by name")
            db.commit()
            stu = conn.fetchall()
            Stu = []
            kemu = [3]
            x = []
            for i in range(len(stu)):
                Stu.append(stu[i][0])
            student, flag = QInputDialog.getItem(self, "学生", "请选择学生：", Stu, 0, False)
            if flag:
                self.statusbar.showMessage('当前显示学生：' + str(student))
                # 创建文件夹
                add = self.osaddress + "/Image/" + str(student)
                if not os.path.exists(add):
                    os.makedirs(add)
                else:
                    shutil.rmtree(self.osaddress + "\\Image\\" + str(student))
                    os.makedirs(add)
                conn.execute("select * from " + student)
                grade = conn.fetchall()
                db.commit()
                conn.close()
                db.close()
                for i in range(len(grade)):
                    x.append(grade[i][0])
                # 读取所选科目
                for i in range(13, 23):
                    if grade[0][i] is not None:
                        kemu.append(i + 1)
                # 绘图
                for i in range(len(kemu)):
                    plt.clf()
                    y = []
                    imageadd = ''
                    km = self.column[kemu[i]]
                    for j in range(len(grade)):
                        y.append(grade[j][kemu[i] - 1])
                    plt.plot(x, y, label=km)
                    plt.xlabel('考试场次')
                    plt.ylabel('排名')
                    plt.title(km)
                    # 显示坐标点的值
                    for a, b in zip(x, y):
                        plt.text(a, b, str(b), ha="center", va="bottom")
                        imageadd = add + "/" + str(km) + ".jpg"
                    plt.savefig(imageadd)
                    self.label[i].setPixmap(QPixmap(imageadd))
        if choose == '帮助':
            QMessageBox.information(None, "帮助", "1.在创建完班级或添加完成绩后需重启软件来加载。\n"
                                                "2.绘图功能所绘的图片全部保存在Image文件夹中，需手动删除\n"
                                                "3.如果遇到导入成绩或添加班级时崩溃，可能是excel文件编码错误，可将.xls文件另存为.xlsx文件来解决问题\n"
                                                "4.添加完班级或成绩后建议刷新一下\n")

    # TableWidght表格设定内容
    def showinfo(self):
        dbaddress = self.combobox.currentText()
        db = sql.connect(self.osaddress + '/AllClass/' + dbaddress)
        conn = db.cursor()
        # 读取表内学生名
        conn.execute("select name from sqlite_master where type='table' order by name")
        db.commit()
        stu = conn.fetchall()
        # 读取所有考试名称
        conn.execute("select filename from " + str(stu[0][0]))
        nameall = conn.fetchall()
        allname = []
        for i in range(len(nameall)):
            allname.append(nameall[i][0])
        self.tableWidget.setRowCount(len(stu) * len(allname))
        for i in range(len(stu)):
            conn.execute("select * from " + str(stu[i][0]))
            result = conn.fetchall()
            row = len(allname)
            vol = len(self.column)
            for j in range(row):
                for k in range(vol):
                    if k == 0:
                        data = QtWidgets.QTableWidgetItem(str(stu[i][0]))
                        self.tableWidget.setItem(i * len(result) + j, k, data)
                    else:
                        data = QtWidgets.QTableWidgetItem(str(result[j][k - 1]))
                        self.tableWidget.setItem(i * len(result) + j, k, data)
        self.tableWidget.setAlternatingRowColors(True)
        conn.close()
        db.close()

    # 导入新成绩
    def AddGrade(self):
        dbaddress, flag = QInputDialog.getItem(self, "班级", "请选择班级：", self.list, 0, False)
        address = self.add.textEdit.toPlainText()
        name = str(self.add.name.text())
        classname = ''
        if address and name and flag:
            # 读取班级
            for i in range(len(dbaddress) - 1, -1, -1):
                if dbaddress[i] == '.':
                    classname = dbaddress[:i]
            AddRess = self.osaddress + r'/AllClass/' + dbaddress
            db = sql.connect(AddRess)
            conn = db.cursor()
            try:
                df = pd.read_excel(address)
            except:
                QMessageBox.critical(None, '错误', '崩溃了.QAQ.')
                QMessageBox.critical(None, '错误', '导入成绩或添加班级时崩溃，可能是excel文件编码错误，可尝试将.xls文件另存为.xlsx文件来解决问题')
            flag1 = False
            flag2 = False
            i = 0
            while not (flag1 and flag2):
                if not np.isnan(df['班级'][i]):
                    if str(df['班级'][i]) == str(classname) and not flag2:
                        flag1 = True
                        conn.execute("insert into " + str(df['姓名'][i]) + "(filename,zongfen,zongmin,yuwen,shuxue,"
                                                                         "yingyu,wuli,huaxue,shengwu,zhengzhi,lishi,"
                                                                         "dili,jishu,yumin,shumin,yinmin,wumin,"
                                                                         "huamin,shengmin,zhengmin,shimin,dimin,"
                                                                         "jimin) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
                                                                         "?,?,?,?,?,?,?,?,?)",
                                     (name, df.at[i, '总分'], df['总名'][i], df['语文'][i], df['数学'][i], df['英语'][i],
                                      df['理赋'][i],
                                      df['化赋'][i], df['生赋'][i], df['政赋'][i], df['史赋'][i], df['地赋'][i], df['技术赋'][i],
                                      df['语名'][i],
                                      df['数名'][i], df['英名'][i], df['理名'][i], df['化名'][i], df['生名'][i], df['政名'][i],
                                      df['史名'][i],
                                      df['地名'][i], df['技术名'][i]))
                        db.commit()
                elif np.isnan(df['班级'][i]) and flag1:
                    flag2 = True
                    conn.close()
                    db.close()
                    self.add.close()
                    QMessageBox.information(None, '提示', '导入完成', QMessageBox.Ok)
                i += 1

        elif not address:
            QMessageBox.critical(None, '错误', '请选择文件', QMessageBox.Ok)
        elif not name:
            QMessageBox.critical(None, '错误', '请输入考试名称', QMessageBox.Ok)
        elif not flag:
            QMessageBox.critical(None, '错误', '请选择数据库', QMessageBox.Ok)

    # 添加班级
    def AddClass(self):
        address = self.add.textEdit.toPlainText()
        name = str(self.add.name.text())
        stuname = []
        type = ''
        if address and not name + '.db' in self.list:
            for i in range(len(address) - 1, -1, -1):  # 获取文件类型
                if address[i] == '.':
                    type = address[i:]
                    break
            if type in ['.xlsx', '.xls']:
                try:
                    df = pd.read_excel(address)
                except:
                    QMessageBox.critical(None, '错误', '崩溃了.QAQ.')
                    QMessageBox.critical(None, '错误', '导入成绩或添加班级时崩溃，可能是excel文件编码错误，可尝试将.xls文件另存为.xlsx文件来解决问题')
                for i in range(len(df)):
                    if str(df.at[i, '班级']) == name:
                        stuname.append(df.at[i, '姓名'])
            # 补充当文件为文本文件时的代码
            db = sql.connect(self.osaddress + '/AllClass/' + name + '.db')
            conn = db.cursor()
            for i in range(len(stuname)):
                conn.execute('create table if not exists ' + stuname[i]
                             + '(filename text,zongfen real,zongmin real,yuwen real,shuxue real,yingyu real,'
                               'wuli real,huaxue real,shengwu real,zhengzhi real,lishi real,dili real,jishu real,'
                               'yumin real,shumin real,yinmin real,wumin real,huamin real,shengmin real,'
                               'zhengmin real,shimin real,dimin real,jimin real)')
                db.commit()
            conn.close()
            db.close()
            self.add.close()
            QMessageBox.information(None, '消息', '创建成功', QMessageBox.Ok)
        elif name + '.db' in self.list:
            QMessageBox.critical(None, '错误', '班级已存在', QMessageBox.Ok)
        else:
            QMessageBox.critical(None, '错误', '请选择文件', QMessageBox.Ok)

    # 选择文件
    def bindList(self):
        from PyQt5.QtWidgets import QFileDialog
        dir = QFileDialog()
        dir.setFileMode(QFileDialog.ExistingFile)
        dir.setDirectory('C:\\')
        dir.setNameFilter('excel文件(*.xlsx *.xls)')
        if dir.exec_():
            address = dir.selectedFiles()
            self.add.textEdit.setText(address[0])


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    mainwindow = MainWindow()
    mainwindow.show()
    sys.exit(app.exec_())
