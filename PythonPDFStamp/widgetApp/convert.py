import pdfplumber
import xlwt

from pathlib import Path

class Convert():
    #给定一个pdf目录，可以把这个pdf转换成xlsx格式的excel文件
    def pdf2excel(self,pdf_path):
        i =0
        workbook =xlwt.Workbook() #创建工作簿
        sheet = workbook.add_sheet('Sheet1') #在工作簿中创建sheet
        pdf_path =pdf_path
        pdf =pdfplumber.open(pdf_path)
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    for j in range(len(row)):
                        sheet.write(i,j,row[j])
                    i +=1
        pdf.close()
        path = Path(pdf_path)
        excel_stem = path.stem #获取pdf文件名不包含后缀
        excel_name =excel_stem+'.xlsx'
        excel_path = path.parent/excel_name
        print(excel_path)
        workbook.save(excel_path)

if __name__ == "__main__":
    convert = Convert()
    convert.pdf2excel('C:\\Users\\郑勋\\Desktop\\2022669301北京君泰通达科技有限公司购销合同1.5.pdf')


