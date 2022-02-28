# -*- coding: utf-8 -*-

import sys
import os
import shutil
import re

from pathlib import Path
from PyQt5.QtWidgets import  QApplication, QWidget

from PyQt5.QtCore import  pyqtSlot,pyqtSignal,Qt,QDir

from PyQt5.QtWidgets import  QFileDialog

#from PyQt5.QtGui import

##from PyQt5.QtSql import 

##from PyQt5.QtMultimedia import

##from PyQt5.QtMultimediaWidgets import


from ui_Widget import Ui_Widget

import pdfStamp


class QmyWidget(QWidget): 
   input_pdf_path =""
   output_pdf_path =""
   def __init__(self, parent=None):
      super().__init__(parent)  #调用父类构造函数，创建窗体
      self.ui=Ui_Widget()       #创建UI对象
      self.ui.setupUi(self)     #构造UI界面


##  ==============自定义功能函数========================


##  ==============event处理函数==========================
        
        
##  ==========由connectSlotsByName()自动连接的槽函数============        
        
        
##  =============自定义槽函数===============================        
   @pyqtSlot()
   def on_btnMoveFile_clicked(self):
      if self.ui.rbtnHikOrder.isChecked():
         input_folder_path =Path('C:\\Users\\郑勋\\Desktop\\海康进货合同\\')
         output_folder_path =Path('H:\\崔向阳坚果云\\进货合同\\0杭州海康智能科技有限公司\\2月\\')
         
      elif self.ui.rbtnHikBorrowOrder.isChecked():
            input_folder_path =Path('C:\\Users\\郑勋\\Desktop\\海康借入合同\\')
            output_folder_path =Path('H:\\崔向阳坚果云\\借入合同\\杭州海康智能科技有限公司\\2月\\')
      
      elif self.ui.rbtnJunLendOrder.isChecked():
            input_folder_path =Path('C:\\Users\\郑勋\\Desktop\\君泰借出合同\\')
            output_folder_path =Path('H:\\崔向阳坚果云\\借出合同\\')
               
      file_list = input_folder_path.glob('*.*')
      for file in file_list:
         file_name = file.name
         chinese_str ="".join(re.findall('[\u4e00-\u9fa5]',file_name))
         company_name =chinese_str.split("借")[0]
         src_file_path = input_folder_path/file_name
         if self.ui.rbtnJunLendOrder.isChecked():
            dst_file_path = output_folder_path/company_name/file_name
            print(dst_file_path)
         else:
            dst_file_path = output_folder_path/file_name
         shutil.move(src_file_path,dst_file_path)


   @pyqtSlot()
   def on_buyPathSetupBtn_clicked(self):
      buyPathStr = QFileDialog.getExistingDirectory(self,'选择采购合同文件夹',os.getcwd())
      self.ui.buyFilePathlLineEdit.setText(buyPathStr)

   @pyqtSlot()
   def on_btnStamp_clicked(self):
      if self.ui.rbtnHikOrder.isChecked():
         input_folder_path =Path('C:\\Users\\郑勋\\Desktop\\海康进货合同\\')
         file_list = input_folder_path.glob('*.pdf*')
         lists =[]
         for file in file_list:
            file_name = file.name
            out_file_name = file_name + "-已盖章.pdf"
            self.input_pdf_path =input_folder_path/file_name
            self.output_pdf_path=input_folder_path/out_file_name
            lists.append(self.input_pdf_path)
            lists.append(self.output_pdf_path)
            # 获取海康进货合同页数
            input_pdf_pages = pdfStamp.get_order_pages(self.input_pdf_path.__str__())
            #watermark是水印文件的路径
            watermark_path = pdfStamp.get_watermark_file(input_pdf_pages)
            pdfStamp.create_watermark(self.input_pdf_path.__str__(),self.output_pdf_path.__str__(),watermark_path)
         print(lists)
      elif self.ui.rbtnHikBorrowOrder.isChecked():
         input_folder_path =Path('C:\\Users\\郑勋\\Desktop\\海康借入合同\\')
         file_list = input_folder_path.glob('*.pdf*')
         lists =[]
         for file in file_list:
            file_name = file.name
            out_file_name = file_name + "-已盖章.pdf"
            self.input_pdf_path =input_folder_path/file_name
            self.output_pdf_path=input_folder_path/out_file_name
            lists.append(self.input_pdf_path)
            lists.append(self.output_pdf_path)
            # 获取海康进货合同页数
            input_pdf_pages = pdfStamp.get_order_pages(self.input_pdf_path.__str__())
            #watermark是水印文件的路径
            watermark_path = pdfStamp.get_watermark_file(input_pdf_pages)
            pdfStamp.create_watermark(self.input_pdf_path.__str__(),self.output_pdf_path.__str__(),watermark_path)
         print(lists)
##  ============窗体测试程序 ================================
if  __name__ == "__main__":        #用于当前窗体测试
   app = QApplication(sys.argv)    #创建GUI应用程序

   form=QmyWidget()                #创建窗体
   form.show()

   sys.exit(app.exec_())
