#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author: 熊书灵
# CreatTime: 2020/4/18
# Email: shulingxiong@163.com


from PySide2.QtWidgets import QApplication, QFileDialog
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile
from PySide2.QtGui import QIcon
from excel2dbc import Excel2Dbc
import os


class MyWindow:

    def __init__(self):
        # 从文件中加载UI定义
        file = QFile('ui/my.ui')
        file.open(QFile.ReadOnly)
        file.close()

        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        self.ui = QUiLoader().load(file)

        self.ui.setWindowTitle('Excel<--->DBC工具')
        # 设置窗口图标
        # self.ui.setWindowIcon(QIcon('./ico/logo.png'))

        # pushButton这个名字和qt设计上的名字要一致
        # 把 pushButton 被 点击（clicked）的信号（signal）连接（connect）到了 handleCalc 这样的一个槽（slot）上
        # 选择文件
        self.ui.pushButton.clicked.connect(self.openfile)

        # 点击转换
        self.ui.pushButton_1.clicked.connect(self.convert)
        self.ui.pushButton_1.setEnabled(False)  # 禁用转换按钮

        self.filepath = []  # 所有文件的绝对路径
        self.path_name = []  # os分割出来的路径和文件名
        self.path = ''  # 路径
        self.file = ''  # 文件名
        self.name = ''  # 无后缀名字
        self.extension = ''  # 后缀
        self.filename = ''  # 传入excel2dbc作为文件名

    def openfile(self):
        self.ui.pushButton_1.setEnabled(True)  # 启用转换按钮

        # 生成文件对话框
        dialog = QFileDialog()
        # 设置文件过滤器，这里是可以打开任何文件（多个）
        dialog.setFileMode(QFileDialog.ExistingFiles)
        # 过滤文件，只显示Excel、DBC文件
        dialog.setNameFilter("*.dbc *.xls *.xlsx")
        # 显示文件模式，这里是详细模式
        dialog.setViewMode(QFileDialog.Detail)
        if dialog.exec_():
            self.filepath = dialog.selectedFiles()  # 返回一个列表, 元素是字符串
            for path in self.filepath:
                # 在编辑框显示路径
                self.ui.plainTextEdit.appendPlainText(path)  # 参数为字符串

    # 选择转换方式
    def convert(self):
        # 分割路径，名称，后缀
        for path in self.filepath:
            self.path_name = os.path.split(path)  # 分割路径和文件, 返回字典类型
            self.path = self.path_name[0]  # 路径, 字符串类型
            self.file = self.path_name[1]  # 文件名, 字符串类型

            self.name = os.path.splitext(self.file)[0]  # 无后缀文件名, 字符串类型
            self.extension = os.path.splitext(self.file)[1]  # 文件后缀，字符串类型
            self.filename = self.path + self.name  # 传入excel2dbc作为文件名

            if self.extension == '.dbc':
                self.ui.plainTextEdit.appendPlainText('DBC2Excel开始转换')
                # 创建dbc2excel对象
                # d2e = Dbc2Excel(path, self.filename)
                self.ui.pushButton_1.setEnabled(False)  # 禁用转换按钮

            elif self.extension == '.xlsx' or self.extension == '.xls':
                self.ui.plainTextEdit.appendPlainText('Excel2DBC开始转换')
                # 创建excel2dbc对象
                e2d = Excel2Dbc(path, self.filename)
                self.ui.pushButton_1.setEnabled(False)  # 禁用转换按钮

            else:
                self.ui.plainTextEdit.appendPlainText('请选择格式正确的文件，本工具支持 .dbc / .xls / .xlsx')
                return


if __name__ == '__main__':
    # 创建一个应用对象
    app = QApplication()
    # 设置主窗口图标
    app.setWindowIcon(QIcon('./ico/logo.png'))

    window = MyWindow()
    # 显示
    window.ui.show()

    # 进入QApplication的事件处理循环，接收用户的输入事件
    app.exec_()
