# -*- coding: utf-8 -*-

from datetime import datetime
import sys
import os
import shutil
import re

from pathlib import Path
from win32com import client

from PyQt5.QtWidgets import QApplication, QWidget,QHeaderView

from PyQt5.QtCore import pyqtSlot, QFileSystemWatcher, pyqtSignal, Qt, QDir

from PyQt5.QtWidgets import QFileDialog, QMessageBox, QStatusBar,QTableWidgetItem

#from PyQt5.QtGui import

##from PyQt5.QtSql import

from ui_Widget import Ui_Widget

import pdfStamp

from wechatAuto import WeChatAuto
from vba import VBA


class QmyWidget(QWidget):
    input_pdf_path = ""
    output_pdf_path = ""

    def __init__(self, parent=None):
        super().__init__(parent)  #调用父类构造函数，创建窗体
        self.ui = Ui_Widget()  #创建UI对象
        self.ui.setupUi(self)  #构造UI界面
        self.statusBar = QStatusBar()
        self.fileWatcher = QFileSystemWatcher()
        self.fileWatcher.directoryChanged.connect(self.do_directoryChanged)
        self.ui.leBorrowFilePath.setText('1111111')
        self.wechat = WeChatAuto()
        self.vba = VBA()
        self.input_folder_path = Path()
        self.output_folder_path = Path()
        self.ui.rbtnHikBorrowOrder.clicked.connect(self.do_show_dir)
        self.ui.rbtnHikOrder.clicked.connect(self.do_show_dir)
        self.ui.rbtnJunLendOrder.clicked.connect(self.do_show_dir)
        self.ui.rbtnJunOrder.clicked.connect(self.do_show_dir)

##  ==============自定义功能函数========================
    #获取目录下最新修改的文件
    def get_new_file(self, folder_path):
        file_list = os.listdir(folder_path)
        print('没有处理的文件列表 \n %s \n' % file_list)
        file_list.sort(key=lambda x: os.path.getmtime(folder_path + '\\' + x))
        file_new = os.path.join(folder_path, file_list[-1])
        print('时间排序后的文件夹列表： \n %s \n' % file_list)
        print('最新文件路径： \n %s' % file_new)
        return file_new

    def get_all_file_dir(self, folder_path):
        file_path_list = []
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            file_path_list.append(file_path)
        return file_path_list

    def set_file_path(self):

        if self.ui.rbtnHikOrder.isChecked():
            self.input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\海康进货合同\\')
            self.output_folder_path = Path(
                'H:\\崔向阳坚果云\\进货合同\\0杭州海康智能科技有限公司\\3月\\')

        elif self.ui.rbtnHikBorrowOrder.isChecked():
            self.input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\海康借入合同\\')
            self.output_folder_path = Path(
                'H:\\崔向阳坚果云\\借入合同\\杭州海康智能科技有限公司\\3月\\')

        elif self.ui.rbtnJunLendOrder.isChecked():
            self.input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\君泰借出合同\\')
            self.output_folder_path = Path('H:\\崔向阳坚果云\\借出合同\\')

        elif self.ui.rbtnJunOrder.isChecked():
            self.input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\销售合同\\')
            self.output_folder_path = Path('H:\\崔向阳坚果云\\销售合同\\')

##  ==============event处理函数==========================

##  ==========由connectSlotsByName()自动连接的槽函数============
#添加销售订单到销售汇总表

    @pyqtSlot()
    def on_btnExcelSaleDetail_clicked(self):
        range_value = self.vba.write_to_sales_detail()
        table_row_count = len(range_value)
        table_column_count = len(range_value[0])
        self.ui.twExcelInfo.setColumnCount(table_column_count)
        self.ui.twExcelInfo.setRowCount(table_row_count)
        self.ui.twExcelInfo.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ui.twExcelInfo.setHorizontalHeaderLabels(
            ['合同编号', '客户名称', '签单日期', '货物名称', '规格型号', '数量', '金额', '合计金额'])
        for row in range(table_row_count):
            for column in range(table_column_count):
                if type(range_value[row][column]) == datetime:
                    range_value[row][column] = range_value[row][column].date().isoformat()
                elif type(range_value[row][column]) == float:
                    range_value[row][column] = str(range_value[row][column])
                new_item = QTableWidgetItem(range_value[row][column])
                #print(new_item)
                self.ui.twExcelInfo.setItem(row,column,new_item)

    #移动文件到微信用户
    @pyqtSlot()
    def on_btnSendFile_clicked(self):
        self.set_file_path()
        if self.ui.rbtnJunOrder.isChecked() and self.ui.rbtnLijie.isChecked():
            file_list = self.input_folder_path.glob('*.pdf')
            self.wechat.findUser('汉子哥')
            for file in file_list:
                #print(file) 这里面是一个str类型的file路径 C:\Users\郑勋\Desktop\销售合同\202203081055-南阳理工学院购销合同.pdf
                #self.wechat.findImg(self.wechat.image_path)
                self.wechat.sendFile(str(file), self.wechat.image_path)

    #添加监控目录
    @pyqtSlot()
    def on_btnAddWatchPath_clicked(self):
        #curDir = QDir.currentPath()
        wechat_file_path = 'C:\\Users\\郑勋\\Documents\\WeChat Files\\q37610672\\FileStorage\\File\\2022-03'
        path = QFileDialog.getExistingDirectory(self, "选择一个要监听的目录",
                                                wechat_file_path,
                                                QFileDialog.ShowDirsOnly)
        self.fileWatcher.addPath(path)

    # 点击李杰radioButton
    @pyqtSlot()
    def on_rbtnLijie_clicked(self):
        self.ui.rbtnHikBorrowOrder.setChecked(False)
        self.ui.rbtnHikOrder.setChecked(False)

    # 点击海康下单群radioButton
    @pyqtSlot()
    def on_rbtnHikGroup_clicked(self):
        self.ui.rbtnJunLendOrder.setChecked(False)
        self.ui.rbtnJunOrder.setChecked(False)

    # 移动文件槽函数
    @pyqtSlot()
    def on_btnMoveFile_clicked(self):
        self.set_file_path()
        file_list = self.input_folder_path.glob('*.*')
        for file in file_list:
            file_name = file.name
            src_file_path = self.input_folder_path / file_name
            chinese_str = "".join(re.findall('[\u4e00-\u9fa5]', file_name))

            if self.ui.rbtnJunLendOrder.isChecked():
                company_name = chinese_str.split("借")[0]
                #判断外借合同客户的目录是否存在，不存在则新建目录后再移动文件
                if not os.path.exists(self.output_folder_path / company_name):
                    os.makedirs(self.output_folder_path / company_name)
                dst_file_path = self.output_folder_path / company_name / file_name
                print(dst_file_path)
            elif self.ui.rbtnJunOrder.isChecked():
                company_name = chinese_str.split("购")[0]
                if not os.path.exists(self.output_folder_path / company_name):
                    os.makedirs(self.output_folder_path / company_name)
                dst_file_path = self.output_folder_path / company_name / file_name
                print(dst_file_path)
            else:
                dst_file_path = self.output_folder_path / file_name
            shutil.move(src_file_path, dst_file_path)

    #
    @pyqtSlot()
    def on_buyPathSetupBtn_clicked(self):
        buyPathStr = QFileDialog.getExistingDirectory(self, '选择采购合同文件夹',
                                                      os.getcwd())
        self.ui.buyFilePathlLineEdit.setText(buyPathStr)

    #给文件盖章槽函数
    @pyqtSlot()
    def on_btnStamp_clicked(self):
        try:
            if self.ui.rbtnHikOrder.isChecked():
                input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\海康进货合同\\')
                file_list = input_folder_path.glob('*.pdf*')
                lists = []
                for file in file_list:
                    file_name = file.name
                    out_file_name = file_name + "-已盖章.pdf"
                    self.input_pdf_path = input_folder_path / file_name
                    self.output_pdf_path = input_folder_path / out_file_name
                    lists.append(self.input_pdf_path)
                    lists.append(self.output_pdf_path)
                    # 获取海康进货合同页数
                    input_pdf_pages = pdfStamp.get_order_pages(
                        self.input_pdf_path.__str__())
                    #watermark是水印文件的路径
                    watermark_path = pdfStamp.get_watermark_file(
                        input_pdf_pages,self.ui)
                    pdfStamp.create_watermark(self.input_pdf_path.__str__(),
                                              self.output_pdf_path.__str__(),
                                              watermark_path)
                print(lists)
            elif self.ui.rbtnHikBorrowOrder.isChecked():
                input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\海康借入合同\\')
                file_list = input_folder_path.glob('*.pdf*')
                lists = []
                for file in file_list:
                    file_name = file.name
                    out_file_name = file_name + "-output.pdf"
                    self.input_pdf_path = input_folder_path / file_name
                    self.output_pdf_path = input_folder_path / out_file_name
                    lists.append(self.input_pdf_path)
                    lists.append(self.output_pdf_path)
                    # 获取海康进货合同页数
                    input_pdf_pages = pdfStamp.get_order_pages(
                        self.input_pdf_path.__str__())
                    #watermark是水印文件的路径
                    watermark_path = pdfStamp.get_watermark_file(
                        input_pdf_pages, self.ui)
                    pdfStamp.create_watermark(self.input_pdf_path.__str__(),
                                              self.output_pdf_path.__str__(),
                                              watermark_path)
                print(lists)
            elif self.ui.rbtnJunOrder.isChecked():
                input_folder_path = Path('C:\\Users\\郑勋\\Desktop\\销售合同\\')
                file_list = input_folder_path.glob('*.*')
                for file in file_list:
                    file_name = file.name
                    file_stem = file.stem
                    out_file_name = file_stem + ".pdf"
                    self.input_pdf_path = input_folder_path / file_name
                    print(self.input_pdf_path)
                    self.output_pdf_path = input_folder_path / out_file_name
                    print(self.output_pdf_path)
                    if file.suffix == ".xlsx":
                        # Open Microsoft Excel
                        excel = client.Dispatch("Excel.Application")
                        #excel.Visible = False #后台运行
                        excel.DisplayAlerts = False  #禁止弹窗
                        # Read Excel File
                        workbook = excel.Workbooks.Open(
                            self.input_pdf_path.__str__())
                        worksheet = workbook.Worksheets[0]
                        #Convert into PDF
                        workbook.ExportAsFixedFormat(
                            0, self.output_pdf_path.__str__())
                        file_pdf = input_folder_path.glob('*.pdf')
                        for file in file_pdf:
                            file_name = file.name
                            self.input_pdf_path = input_folder_path / file_name
                            self.output_pdf_path = input_folder_path / file_name
                            watermark_path = "K:\\GithubCode\\juntevision\\PythonPDFStamp\\pdf\\盖君泰的合同1页版本水印.pdf"
                            pdfStamp.create_watermark(
                                self.input_pdf_path.__str__(),
                                self.output_pdf_path.__str__(), watermark_path)
                        # work_sheets.Close()
                        excel.DisplayAlerts = True
                    elif file.suffix == '.pdf':
                        watermark_path = "K:\\GithubCode\\juntevision\\PythonPDFStamp\\pdf\\盖君泰的合同1页版本水印.pdf"
                        pdfStamp.create_watermark(
                            self.input_pdf_path.__str__(),
                            self.output_pdf_path.__str__(), watermark_path)
        except Exception as e:
            QMessageBox.information(self, '文件盖章出错', e.__str__())

##  =============自定义槽函数===============================

    def do_show_dir(self):
        self.ui.lwFileWatcher.clear()
        self.set_file_path()
        file_list = self.get_all_file_dir(self.input_folder_path)
        print(file_list)
        print(str(self.input_folder_path))
        self.ui.lwFileWatcher.addItems(file_list)

    # 目录内增加文件的时候，调用的槽函数
    def do_directoryChanged(self, path):
        self.ui.lwFileWatcher.clear()
        file_names = os.listdir(path)
        print(file_names)
        #for file_name in file_names:
        #    self.ui.lwFileWatcher.addItem(file_name)
        new_added_file = self.get_new_file(path)
        self.ui.lwFileWatcher.addItem(new_added_file)


##  ============窗体测试程序 ================================
if __name__ == "__main__":  #用于当前窗体测试
    app = QApplication(sys.argv)  #创建GUI应用程序

    form = QmyWidget()  #创建窗体
    form.show()

    sys.exit(app.exec_())
