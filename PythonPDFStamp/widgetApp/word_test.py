
from docx import Document


class Word():
    def get_sale_order_word(self,path):
        doc = Document(path)
        tables = doc.tables
        table_one =tables[0]
        #print(type(table_one._cells))
        sale_order_info_list=[]
        sale_order_info_all=[]
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
            sale_order_info_dic={}
            sale_order_info_dic['购货单位']=sale_order_info_list[0]
            sale_order_info_dic['合同编号']=sale_order_info_list[1]
            sale_order_info_dic['签订日期']=table_three._cells[-1].text.split('：')[1]
            sale_order_info_dic['商品名称'] =row.cells[1].text
            sale_order_info_dic['型号'] =row.cells[2].text
            sale_order_info_dic['数量'] =row.cells[4].text
            sale_order_info_dic['单价']= row.cells[5].text
            sale_order_info_all.append(sale_order_info_dic)
            #print(sale_order_info_dic)
        print(sale_order_info_all)
        return sale_order_info_all

if __name__ == '__main__':
    path = 'H:\\崔向阳坚果云\\销售合同\\济南微纳颗粒仪器股份有限公司\\1月\\202201041001-济南微纳颗粒购销合同.docx'
    word = Word()
    word.get_sale_order_word(path)