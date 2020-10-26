# -*- coding: utf-8 -*-
import sys
import xdrlib
import xlrd
import os
import shutil
##########################################################
reload(sys)
sys.setdefaultencoding('utf-8')

##########################################################
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

##########################################################
def main(argv):
    data = open_excel(sys.argv[1])
    if data:
        table = data.sheets()[0]
        colnames = table.row_values(0) #第一行数据
        colKeys = table.col_values(0) #第一列key数据
        nrows = len(colKeys) #总行数
        ncols = len(colnames) #总列数

        #语言列表
        languageList = []
        #国际化文件名列表
        fileNameList = []
        #国际化常量量列表
        localizableConstantList =[]
        #国际化变量列表
        localizableStringList = []

        parent = os.path.dirname(os.path.realpath(__file__))
        dirPath =  parent + '/localize'
        if not os.path.exists(dirPath):
            os.mkdir(dirPath)

        for indexCol in range(1,ncols):
            list = []
            list2 = []
            list3 = []
            colValues = table.col_values(indexCol)
            fileNameList.append(colValues[0] + '.strings')
            for indexRow in range(1,nrows):
                
                value = colValues[indexRow]
                # if (len(value)==0):
                    # value = colValues_English[indexRow]

                key = colKeys[indexRow]
                #key2 （key的首字母强制改为大写。比如key:warmlyTips，通过这个方法格式化为：Localizable_WarmlyTips）
                key2 = 'Localizable_' + key.replace(key[0],key[0].capitalize())
                
                
                keyValue = '"' + key2 + '"' + ' = ' + '"' + value + '"' + ';\n'
                list.append(keyValue)

               
                keyValue2 = 'static let ' + key2 +  ' = ' + '"' + key2 + '"' + '\n'
                list2.append(keyValue2)


                keyValue3 = 'static var ' + key +  ' : String { return NSLocalizedString(LocalizableConstants.' + key2 + ',comment:  "") }' + '\n'
                list3.append(keyValue3)

            languageList.append(''.join(list))
            localizableConstantList.append(''.join(list2))
            localizableStringList.append(''.join(list3))

        for index in range(len(languageList)):
            fileName = dirPath + '/' + fileNameList[index]
            os.system(r'touch %s' % fileName)
            
            fp = open(fileName,'wb+')
            fp.write(languageList[index])
            fp.close()

        for index in range(len(localizableConstantList)):
            fileName = dirPath + '/' + 'LocalizableConstants.swift'
            os.system(r'touch %s' % fileName)

            fp = open(fileName, 'wb+')
            fp.write(localizableConstantList[index])
            fp.close()

        for index in range(len(localizableStringList)):
            fileName = dirPath + '/' + 'String+Globalization.swift'
            os.system(r'touch %s' % fileName)

            fp = open(fileName, 'wb+')
            fp.write(localizableStringList[index])
            fp.close()


    else :
                print "can not open file"

if __name__=="__main__":
    main(sys.argv[1])
