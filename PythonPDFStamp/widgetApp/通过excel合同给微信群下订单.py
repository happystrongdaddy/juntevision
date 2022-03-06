import xlwings as xw
from xlwings.main import Sheet
import pyautogui
import pyperclip
import time

# 把内嵌list合成一般list
outputCells = []


def removeNestings(nestList):
    for i in nestList:
        if type(i) == list:
            removeNestings(i)
        else:
            outputCells.append(i)


app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False
path = r"C:\Users\郑勋\Desktop\合同\202202252010-博科视（苏州）技术有限公司-购销合同.xlsx"
wb = app.books.open(path)
sheet = wb.sheets['Sheet1']
# 获取有效的excel的区域
info = sheet.used_range
# 获取有效区域的行列数
nrows = info.last_cell.row
ncolumns = info.last_cell.column
# 获取excel表单内的所有内容
allcells = sheet.range((1, 1), (nrows, ncolumns)).value
# 把内嵌list合成普通list
removeNestings(allcells)
# 查找出地址信息
addressStr = ""
for cell in outputCells:
    if cell == None:
        continue
    elif not isinstance(cell, str):
        continue
    elif "交货地点" in cell:
        addressStr = cell
addressStr = addressStr.lstrip()
addressStr = addressStr.split('.', 1)
# 获取订购型号和数量
typeCells = sheet.range("D11:D20").value
numCells = sheet.range("F11:F20").value
# 合成需要输出到微信的字符串
outputBeginStr = "@李浩  HIKROBOT 麻烦出一个购销合同 "
outputString = ""
for (type, num) in zip(typeCells, numCells):
    if(type != None or num != None):
        outputString = outputString + type+" 数量"+str(int(num))+"个 "
    else:
        break
outputFinalStr = outputBeginStr+outputString + addressStr[1]
print(outputFinalStr)
# 打开微信
pyautogui.hotkey('ctrl', 'alt', 'w')
time.sleep(1)

# 搜索群聊名称--君泰-海康商务群-下单
pyautogui.hotkey("ctrl", "f")
pyperclip.copy("君泰-海康商务群-下单")
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
pyautogui.hotkey('Enter')
time.sleep(1)
# 拷贝要发送的内容
pyperclip.copy(outputFinalStr)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)

wb.close()
app.quit()
