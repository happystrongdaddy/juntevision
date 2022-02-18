import os

#this demo gives the order number from the path string
filePath = "C:/Users/郑勋/桌面/2022629247北京君泰通达科技有限公司购销合同12.31-北航.pdf"
pathNameAndExtension = os.path.splitext(filePath)
print('pathNameAndExtension的类型是%s' % type(pathNameAndExtension))
print('pathNameAndExtension is %s' % str(pathNameAndExtension))
pathName = os.path.splitext(filePath)[0]
print(pathName)
print(type(pathName))
fileName = pathName.split('/')[-1]
print(fileName)
print(type(fileName))
fileName = fileName[0:10]
print(fileName)
print(type(fileName))
print(int(fileName))
