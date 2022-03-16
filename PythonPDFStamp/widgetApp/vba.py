import xlwings as xw
import time

class VBA():
    #20220309 还没测试好
    def write_to_sales_detail(self):
        app =xw.App(visible=False,add_book=False)
        app.screen_updating = False
        file_path ='H:\\崔向阳坚果云\\2022年销售业绩明细单.xlsm'
        wb = app.books.open(file_path)
        sheet =wb.sheets[0]
        # 工作表sheet中有数据区域最大的行数
        row_last_before = sheet.used_range.last_cell.row+1
        print(row_last_before)
        column_last_before = sheet.used_range.last_cell.column
        print(column_last_before)
        marco = wb.macro('getDataFromExcel')
        marco()
        row_last_after = sheet.used_range.last_cell.row
        print(row_last_after)
        new_range =sheet.range((row_last_before,1),(row_last_after,column_last_before)).options(ndim=2).value
        print(new_range)
        #print(type(new_range[0][2]))
        #print(new_range[0][2])
        time.sleep(1)
        app.screen_updating = True
        wb.save()
        wb.close()
        app.quit()
        app.kill()
        return new_range
        
if __name__ == '__main__':
    vba = VBA()
    vba.write_to_sales_detail()