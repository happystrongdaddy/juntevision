import os

import pdfplumber
import pymysql

from pathlib import Path
from docx import Document


class Juntvision():
    #把每一行的采购信息，汇总添加到这个list中
    buy_order_list = []

    #获取海康采购订单的详细信息
    def get_buy_order_info(self, buy_order_path):
        pdf = pdfplumber.open(buy_order_path)
        pdf_page_one = pdf.pages[0]
        page_one_text = pdf_page_one.extract_text()
        page_text_list = page_one_text.split()
        #获取海康采购合同订购单号
        buy_order_id = page_text_list[8].split('：')[1]

        #获取海康采购合同日期
        buy_date = page_text_list[9].split('：')[1]

        #获取海康公司名称
        supplier_name = page_text_list[14].split('：')[1]

        #截取订购明细信息表
        page_one_table = pdf_page_one.extract_table()[1:-1]
        #print(page_one_table)

        for product in page_one_table:
            #用于储存采购的每一行信息，然后插入数据库中
            product_info_dic = {}
            product_info_dic['合同编号'] = buy_order_id
            product_info_dic['签订日期'] = buy_date
            product_info_dic['供货方'] = supplier_name
            #print(product)
            product_name = product[1]
            product_info_dic['商品名称'] = product_name
            #print(product_name)
            product_model = product[2]
            if '\n' in product_model:
                product_model = product_model.replace('\n', '')
            product_info_dic['型号'] = product_model
            #print(product_model)
            quantity = product[3]
            product_info_dic['数量'] = quantity
            #print(quantity)
            unit_price = product[4]
            product_info_dic['单价'] = unit_price
            #print(unit_price)
            self.buy_order_list.append(product_info_dic)
            #print(product_info_dic)
        print(len(self.buy_order_list))
        print(self.buy_order_list)
        return self.buy_order_list

    #获取君泰销售订单明细
    def get_sale_order_info(self, path):
        doc = Document(path)
        tables = doc.tables
        table_one = tables[0]
        #print(type(table_one._cells))
        sale_order_info_list = []
        sale_order_info_all = []
        for cell in table_one._cells:
            # print(type(table_one._cells))
            #sale_order_info_dic={}
            sale_order_info_list.append(cell.text.split('：')[1])
            #print(cell.text)
        #获取购货单位和合同编号
        sale_order_info_list = sale_order_info_list[0:2]
        #print(sale_order_info_list)
        table_three = tables[2]
        #获取合同签订日期
        #print(table_three._cells[-1].text)
        sale_order_info_list.append(table_three._cells[-1].text.split('：')[1])
        #print(sale_order_info_list)
        table_two = tables[1]

        for row in table_two.rows[1:-1]:
            sale_order_info_dic = {}
            sale_order_info_dic['购货单位'] = sale_order_info_list[0]
            sale_order_info_dic['合同编号'] = sale_order_info_list[1]
            sale_order_info_dic['签订日期'] = table_three._cells[-1].text.split(
                '：')[1]
            sale_order_info_dic['商品名称'] = row.cells[1].text
            sale_order_info_dic['型号'] = row.cells[2].text
            sale_order_info_dic['数量'] = row.cells[4].text
            sale_order_info_dic['单价'] = row.cells[5].text
            sale_order_info_all.append(sale_order_info_dic)
            #print(sale_order_info_dic)
        #print(sale_order_info_all)
        return sale_order_info_all

    #把海康采购的订单数据录入到数据库中
    def insert_buy_order(self, buy_order_list):
        connection = pymysql.connect(host='localhost',
                                     user='root',
                                     passwd='11251125',
                                     database='juntevision')
        cursor = connection.cursor()
        for product_item in buy_order_list:
            insert_sql="insert into buydetail(buyOrderID, buyDate, supplierName, productName, \
                productModel, quantity, unitPrice) values('%s','%s','%s','%s','%s','%d','%f')"                                                                                               % \
                    (product_item['合同编号'],product_item['签订日期'],product_item['供货方'],product_item['商品名称'], \
                        product_item['型号'],int(product_item['数量']),float(product_item['单价']))
            try:
                cursor.execute(insert_sql)
                connection.commit()
            except:
                connection.rollback()
                print("插入数据异常")
        connection.close()

    def insert_saledetail_table(self, sale_order_list):
        connection = pymysql.connect(host='localhost',
                                     user='root',
                                     passwd='11251125',
                                     database='juntevision')
        cursor = connection.cursor()
        for product_item in sale_order_list:
            print(product_item)
            insert_sql="insert into saledetail(saleOrderID, customerName, saleDate, productName, productModel, \
                 quantity,unitPrice) values('%s','%s','%s','%s','%s','%d','%f')" % \
                    (product_item['合同编号'],product_item['购货单位'],product_item['签订日期'],product_item['商品名称'], \
                        product_item['型号'],int(product_item['数量']),float(product_item['单价']))
            try:
                cursor.execute(insert_sql)
                connection.commit()
            except pymysql.Error as e:
                connection.rollback()
                print("插入数据异常")
                print(e.args)
        connection.close()

    #把海康合同中的产品信息录入到数据库中
    #buy_order_list是从海康订购合同中提取的商品信息
    def insert_product_table(self, buy_order_list):
        connection = pymysql.connect(host='localhost',
                                     user='root',
                                     passwd='11251125',
                                     database='juntevision')
        cursor = connection.cursor()
        for product_item in buy_order_list:
            select_sql = "select productModel from product where productModel = '%s'" % product_item[
                '型号']
            result = cursor.execute(select_sql)
            print(result)
            if result == 0:
                insert_sql="insert into product(productName, productModel, productBuyPrice, supplierName \
                    ) values('%s','%s','%f','%s')"                                                   % \
                        (product_item['商品名称'],product_item['型号'],float(product_item['单价']),product_item['供货方'])
                try:
                    cursor.execute(insert_sql)
                    connection.commit()
                except:
                    connection.rollback()
                    print("插入数据异常")
        connection.close()

    #修改添加完后的采购订单文件名，后面加上mysql字符
    def rename_buy_order(self, buy_order_path):
        #file_name = os.path.basename(buy_order_path)
        path = Path(buy_order_path)
        file_stem = path.stem
        file_new_stem = file_stem + '-mysql'
        file_new_name = str(path.parent / file_new_stem) + path.suffix
        print(file_new_name)
        path.rename(file_new_name)


if __name__ == "__main__":

    path = 'H:\\崔向阳坚果云\\销售合同\\济南微纳颗粒仪器股份有限公司\\1月\\202201041001-济南微纳颗粒购销合同.docx'
    juntObj = Juntvision()
    sale_order_list = juntObj.get_sale_order_info(path)
    juntObj.insert_saledetail_table(sale_order_list)

    # folder_path ='H:\\崔向阳坚果云\\进货合同\\0杭州海康智能科技有限公司\\1月\\'
    # file_path_list =[]
    # for file in os.listdir(folder_path):
    #     file_path = os.path.join(folder_path,file)
    #     file_path_list.append(file_path)
    # print(file_path_list)
    # pdf_path ='C:\\Users\\郑勋\\Desktop\\2022669301北京君泰通达科技有限公司购销合同1.5.pdf'
    # juntObj =Juntvision()
    # #juntObj.rename_buy_order(pdf_path)
    # buy_order_list=juntObj.get_buy_order_info(pdf_path)
    # #juntObj.insert_buy_order(buy_order_list)
    # juntObj.insert_buy_order_product(buy_order_list)