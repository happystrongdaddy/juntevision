import pdfplumber
import pymysql

from pathlib import Path
from convert import Convert


class Juntvision():
    buy_order_list =[]
    product_info_dic={}
    #获取海康采购订单的详细信息
    def get_buy_order_info(self,buy_order_path):        
        pdf =pdfplumber.open(buy_order_path)
        pdf_page_one = pdf.pages[0]
        page_one_text = pdf_page_one.extract_text()
        page_text_list =page_one_text.split()
        #获取海康采购合同订购单号
        buy_order_id = page_text_list[8].split('：')[1]
        self.product_info_dic['合同编号']=buy_order_id
        #获取海康采购合同日期
        buy_date = page_text_list[9].split('：')[1]
        self.product_info_dic['签订日期']=buy_date
        #获取海康公司名称
        supplier_name = page_text_list[14].split('：')[1]
        self.product_info_dic['供货方']=supplier_name
        #截取订购明细信息表
        page_one_table =pdf_page_one.extract_table()[1:-1]
        for product in page_one_table:
            product_name =product[1]
            self.product_info_dic['商品名称']=product_name
            #print(product_name)
            product_model =product[2]
            if '\n' in product_model:
                product_model =product_model.replace('\n','')
            self.product_info_dic['型号']=product_model
            #print(product_model)
            quantity = product[3]
            self.product_info_dic['数量']=quantity
            #print(quantity)
            unit_price = product[4]
            self.product_info_dic['单价']=unit_price
            #print(unit_price)
            self.buy_order_list.append(self.product_info_dic)
        print(len(self.buy_order_list))
        print(self.buy_order_list)
        return self.buy_order_list

    #把海康采购的订单数据录入到数据库中    
    def insert_buy_order(self,buy_order_list):
        connection = pymysql.connect(host='localhost',user='root',passwd='11251125',database='juntevision')
        cursor = connection .cursor()
        for product_item in buy_order_list:
            insert_sql="insert into buydetail(buyOrderID, buyDate, supplierName, productName, \
                productModel, quantity, unitPrice) values('%s','%s','%s','%s','%s','%d','%f')" % \
                    (product_item['合同编号'],product_item['签订日期'],product_item['供货方'],product_item['商品名称'], \
                        product_item['型号'],int(product_item['数量']),float(product_item['单价']))
            try:
                cursor.execute(insert_sql)
                connection.commit() 
            except:
                connection.rollback()
                print("插入数据异常")           
        connection.close()

if __name__ == "__main__":
    pdf_path ='C:\\Users\\郑勋\\Desktop\\2022669301北京君泰通达科技有限公司购销合同1.5.pdf'
    juntObj =Juntvision()
    buy_order_list=juntObj.get_buy_order_info(pdf_path)
    juntObj.insert_buy_order(buy_order_list)