import os
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QInputDialog
from PyQt5.QtWidgets import QMessageBox
from Mainwindow import Ui_MainWindow
import Add
import Rewarding
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
        self.reward = QtWidgets.QAction("奖励")
        self.toolBar.addActions(
            [self.addgrade, self.delgrade, self.delclass, self.shuaxin, self.huitu, self.help, self.reward])
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
                dirs = self.osaddress + '/AllClass/' + dbaddress
                db = sql.connect(dirs)
                conn = db.cursor()
                # 读取表名
                stu = self.Selecting_Tablename(dirs)
                conn.execute("select filename from " + str(stu[0]))
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
                            conn.execute("delete from ? where filename=?", (str(stu[i]), testname,))
                            db.commit()
                        conn.close()
                        db.close()
                        QMessageBox.information(None, '消息', '删除成功')
        if choose == '刷新':
            # QMessageBox.information(None, '提示', '功能尚未完成')
            self.list = os.listdir(self.osaddress + '/AllClass')
            # self.combobox.clear()
            self.combobox.addItems(self.list)
            # self.toolBar.addActions([self.reward])
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
            dirs = self.osaddress + '/AllClass/' + dbaddress
            db = sql.connect(dirs)
            conn = db.cursor()
            # 读取表内学生名
            stu = self.Selecting_Tablename(dirs)
            kemu = [3]
            x = []
            student, flag = QInputDialog.getItem(self, "学生", "请选择学生：", stu, 0, False)
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

        if choose == '奖励':
            from PyQt5.QtWidgets import QInputDialog
            self.reward = Rewarding.Ui_Form()
            self.reward.paimin.clear()
            self.reward.jinbu.clear()
            dbaddress = self.combobox.currentText()
            dirs = self.osaddress + '/AllClass/' + dbaddress
            db = sql.connect(dirs)
            conn = db.cursor()

            # 读取表内学生名
            stu = self.Selecting_Tablename(dirs)
            conn.execute("select filename from %s" % str(stu[0]))
            test1 = conn.fetchall()
            test2 = []
            for i in range(len(test1)):
                test2.append(test1[i][0])
            test, flag = QInputDialog.getItem(self, "选择考试", "请选择本次考试：", test2, 0, False)
            pre_test, flag1 = QInputDialog.getItem(self, "选择考试", "请选择前一次考试：", test2, 0, False)
            if test and pre_test and flag and flag1:
                list_test = []
                list_pre_test = []
                col = ['filename', 'zongfen', 'zongmin', 'yuwen', 'shuxue', 'yingyu', 'wuli', 'huaxue', 'shengwu',
                       'zhengzhi', 'lishi', 'dili', 'jishu', 'yumin', 'shumin', 'yinmin', 'wumin', 'huamin', 'shengmin',
                       'zhengmin', 'shimin', 'dimin', 'jimin']
                # 将所选两次考试成绩筛选出来
                for i in range(len(stu)):
                    conn.execute("select * from " + str(stu[i]) + " where filename = '" + test + "'")
                    x = conn.fetchall()
                    if not x == []:
                        list_test.append(x[0])
                    else:
                        list_test.append(x)
                    conn.execute("select * from " + str(stu[i]) + " where filename = '" + pre_test + "'")
                    x = conn.fetchall()
                    if not x == []:
                        list_pre_test.append(x[0])
                    else:
                        list_pre_test.append(x)
                conn.close()
                db.close()
                # 将考试成绩转为DataFame便于排序
                pd_test = pd.DataFrame(list_test, columns=col)
                pd_pre_test = pd.DataFrame(list_pre_test, columns=col)
                pd_stu = pd.DataFrame(stu, columns=["姓名"])
                pd_test = pd.concat([pd_stu, pd_test], axis=1)
                pd_pre_test = pd.concat([pd_stu, pd_pre_test], axis=1)
                paimin = ['zongfen', 'yuwen', 'shuxue', 'yingyu', 'wuli', 'huaxue', 'shengwu',
                          'zhengzhi', 'lishi', 'dili', 'jishu']
                fanyi = {'zongfen': '总分', 'yuwen': "语文", 'shuxue': "数学", 'yingyu': "英语", 'wuli': "物理", 'huaxue': "化学",
                         'shengwu': "生物",
                         'zhengzhi': "政治", 'lishi': "历史", 'dili': "地理", 'jishu': "技术"}
                # 每科排名
                self.reward.paimin.append("考试" + text + "的各科排名：")
                for i in range(len(paimin)):
                    pd_pai = pd_test.sort_values(paimin[i], ascending=False)
                    pd_pai = pd.concat([pd_pai["姓名"], pd_pai[paimin[i]]], axis=1)
                    j = 1
                    rs = 3
                    k = 0
                    pai = pd_pai.values.tolist()
                    imax = pai[0][1]
                    # 将前几名存于people中
                    people = []
                    if paimin[i] == 'zongfen':
                        rs = 10
                    while j < rs and pai[0][1] is not None:
                        if pai[k][1] == imax:
                            people.append(pai[k])
                        elif pai[k][1] < imax:
                            j += 1
                            people.append(pai[k])
                        k += 1
                    # 将排名放入GUI视图中
                    if people:
                        self.reward.paimin.append(fanyi[paimin[i]] + "的排名为:")
                        for j in range(len(people)):
                            self.reward.paimin.append(str(j + 1) + " " + str(people[j][0]) + " " + str(people[j][1]))
                        self.reward.paimin.append("--------------------------------------------------------")
                # 进步
                jb = []
                for i in range(len(stu)):
                    a = pd_pre_test.at[i, "zongmin"] - pd_test.at[i, "zongmin"]
                    jb.append(a)
                pd_jb = pd.DataFrame(jb, columns=["进步名次"])
                pd_jb = pd.concat([pd_stu, pd_jb], axis=1)
                pd_jb = pd_jb.sort_values("进步名次", ascending=False)
                list_jb = pd_jb.values.tolist()
                self.reward.jinbu.append("考试" + test + "相较于考试" + pre_test + ":")
                for i in range(len(stu)):

                    if list_jb[i][1] >= 50:
                        self.reward.jinbu.append(
                            str(i + 1) + " " + str(list_jb[i][0]) + " " + str(list_jb[i][1]))
                    else:
                        break
                self.reward.show()
            elif not (test and pre_test):
                QMessageBox.information(None, "提示", '请将考试选择完整。')

    # TableWidght表格设定内容
    def showinfo(self):
        dbaddress = self.combobox.currentText()
        dirs = self.osaddress + '/AllClass/' + dbaddress
        db = sql.connect(dirs)
        conn = db.cursor()
        # 读取表内学生名
        stu = self.Selecting_Tablename(dirs)
        # 读取所有考试名称
        conn.execute("select filename from " + str(stu[0]))
        nameall = conn.fetchall()
        allname = []
        for i in range(len(nameall)):
            allname.append(nameall[i][0])

        self.tableWidget.setRowCount(len(stu) * len(allname))
        for i in range(len(stu)):
            conn.execute("select * from " + str(stu[i]))
            result = conn.fetchall()
            conn.execute("select filename from " + str(stu[i]))
            row1 = conn.fetchall()
            row2 = []
            for j in range(len(row1)):
                row2.append(row1[j][0])
            row = len(row2)
            vol = len(self.column)
            for j in range(row):
                for k in range(vol):
                    if k == 0:
                        data = QtWidgets.QTableWidgetItem(str(stu[i]))
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
            address1 = self.osaddress + r'/AllClass/' + dbaddress
            db = sql.connect(address1)
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
                    QMessageBox.information(None, '提示', '导入完成,:)', QMessageBox.Ok)
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
        types = ''
        if address and not name + '.db' in self.list:
            for i in range(len(address) - 1, -1, -1):  # 获取文件类型
                if address[i] == '.':
                    types = address[i:]
                    break
            if types in ['.xlsx', '.xls']:
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
        dirs = QFileDialog()
        dirs.setFileMode(QFileDialog.ExistingFile)
        dirs.setDirectory('C:\\')
        dirs.setNameFilter('excel文件(*.xlsx *.xls)')
        if dirs.exec_():
            address = dirs.selectedFiles()
            self.add.textEdit.setText(address[0])

    def Selecting_Tablename(self, address):
        db = sql.connect(address)
        cur = db.cursor()
        cur.execute("select name from sqlite_master where type='table' order by name")
        a = cur.fetchall()
        b = []
        for i in range(len(a)):
            b.append(a[i][0])
        return b


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    mainwindow = MainWindow()
    mainwindow.show()
    sys.exit(app.exec_())
