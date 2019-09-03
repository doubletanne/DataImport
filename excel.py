import os
import re
import xlrd
import time
import shutil
import sys

from indi import *
from excelUtil import *
from entity import *

class Excel(object):
    #读取固定目录
    def read_excel(self,path,cycleId):
        for fileName in os.listdir(path):
            #只读取excel文件
            if (os.path.splitext(fileName)[1] == '.xls' or os.path.splitext(fileName)[1] == '.xlsx'):
                #取出不带文件后缀的文件名
                str = os.path.splitext(fileName)[-2]
                #取出不带账期的文件名部分
                #正则匹配不带数字的文件名，并取出最后一个下划线
                file_cn_name = re.search(r'\D+',str).group()[:-1]
                # 取出文件名的账期
                # 正则匹配文件名中的数字
                excelCycle = re.search(r'\d+', str).group()
                # 解析指定账期数据
                if excelCycle == cycleId:
                    year = excelCycle[0:4]
                    #判断在数据库表中是否配置该文件
                    excelUtil = ExcelUtil()
                    fileCfg = excelUtil.get_excel_cfg(file_cn_name,year)
                    print(fileCfg)
                    # 文件频率，是按年给，还是按月给
                    file_freq = fileCfg[0][1]
                    # 是否需要解析，如果是直接读取到数据库，则不需要解析，需要对每个单元格数据进行处理的，需要解析
                    need_parse = fileCfg[0][2]
                    if len(fileCfg) > 0:
                        # 读取excel,获取sheet页
                        book_path = os.path.join(path,fileName)
                        print(book_path)
                        with xlrd.open_workbook(book_path) as workbook:
                            sheetBooks = workbook.sheet_names()
                            for sheetName in sheetBooks:
                                #sheetName的匹配年份处理，用于针对sheetName中的年份变量
                                #需要两个变量，一个变量用于存储原变量，一个用于存储数据库变量
                                if sheetName[0:4] == year:
                                    sheetDBName = sheetName[4:]
                                else:
                                    sheetDBName = sheetName
                                # 判断sheet名称是否配置
                                isSheetExist = excelUtil.is_sheet_exist(file_cn_name, sheetDBName, year)
                                if isSheetExist:
                                    # 将读取的记录插入到文件读取历史表
                                    insert_dt = self.timeFormat(self.getTime())
                                    # 写读取日志表
                                    excelUtil.insert_file_read_log(fileName, sheetName, file_cn_name, excelCycle, insert_dt)
                                    # 解析excel
                                    # 在excel_sheet_cfg配置表找到该sheet页的配置信息
                                    print("获取 %s 的配置信息..." % sheetDBName)
                                    sheetCfgs = excelUtil.get_excel_sheet_cfg(file_cn_name, sheetDBName, year)
                                    if sheetCfgs:
                                        for cfg in sheetCfgs:
                                            table = workbook.sheet_by_name(sheetName)
                                            # 财务数据_欠费详细_企业 只需要读取，不需要解析
                                            if need_parse == '0':
                                                self.read_cust_owe_dtl(file_cn_name, sheetDBName, cfg, table,excelCycle)
                                            elif need_parse == '1':
                                                # 读首行，行数据写入org_info
                                                self.get_first_line(file_cn_name, sheetDBName, excelCycle, cfg, table)
                                                # 读首列，列数据写入indi_node
                                                self.get_first_col(file_cn_name, sheetDBName, excelCycle, cfg, table)
                                                # 首列构造层级关系
                                                indi = IndiTree()
                                                indi.gen_indi_tree(file_cn_name, sheetDBName,excelCycle)
                                                # 读取指标数据部分
                                                self.read_excel_data(file_cn_name,sheetDBName,table,excelCycle,cfg[6],cfg[11],cfg[12])


                                else:
                                    print("没有找到 %s %s 的配置信息！" % (fileName, sheetDBName))
                        # 移动excel文件到日期目录
                        self.mv_readed_file(fileName,path)
                    else:
                        print("没有找到 %s 的配置信息！" % fileName)

    def mv_readed_file(self,fileName,path):
        # 判断目录是否存在,不存在则创建日期目录
        dirName = os.path.join(path,str(time.strftime("%Y%m%d", time.localtime(time.time()))))
        print('dirName:' + dirName)
        srcFile = os.path.join(path,fileName)
        targetFile = os.path.join(dirName,fileName)
        if not os.path.exists(dirName):
            print("创建目录 %s " % dirName)
            os.makedirs(dirName)
        # 移动文件
        if os.path.exists(targetFile):
            print('%s 已经存在，删除文件' % targetFile)
            os.remove(targetFile)
        print("移动文件 %s 至目录 %s " % (srcFile,targetFile))
        shutil.move(srcFile,dirName)



    def read_excel_data(self,file_cn_name,sheetName,table,excelCycle,sheetUnit,rowStart,rowEnd):
        #  读取指标数值
        #  1. 删除indicator_data表中kpi_id存在于indi_tree中的数据
        #  2. 按file_nm,sheet_nm 读取org_info表的机构信息
        #  3. 根据机构信息名称获取标准机构信息名称，和机构代码
        #  4. 循环机构信息，获取机构的单元格的列值
        #  5. 读取indi_node表，按file_nm,sheet_nm获取指标信息，获取指标的单元各的行值
        #  6. 获取sheet页的数值单位，换算为元
        #  6. 定位到具体的单元格之后，按照行列交叉取到单元格数据，写表
        excelUtil = ExcelUtil()
        excelUtil.delete_indicator_data_exists(file_cn_name,sheetName,excelCycle,rowStart,rowEnd)

        orgExcelCell = ''
        orgStdName = ''
        orgStdCode = ''

        orgInfo = excelUtil.get_org_names(file_cn_name,sheetName,excelCycle)
        # 指标树没有年份，不同年份不同月份的同个sheet页的同一个名称为一个指标
        indiNode = excelUtil.get_indi_node_indi_id(file_cn_name,sheetName,excelCycle,rowStart,rowEnd)
        if orgInfo:
            for org in orgInfo:
                orgName = org[0]
                orgExcelCell = int(org[1])
                orgStdInfo = self.get_std_org_info(orgName)
                if orgStdInfo:
                    orgStdCode = orgStdInfo[0]
                    orgStdName = orgStdInfo[1]
                    if indiNode:
                        for indi in indiNode:
                            indiId = indi[0]
                            indiName = indi[1]
                            indiExcelCell = indi[2]
                            cellData = None
                            if indiExcelCell != None and indiExcelCell != '':
                                indiExcelCell = int(indi[2])
                                cellData = table.cell_value(indiExcelCell,orgExcelCell)
                            if cellData != None and cellData != '':
                                value1 = self.get_std_num(cellData,sheetUnit)
                                value2 = float(cellData)
                            else:
                                value1 = None
                                value2 = None
                            createTime = self.timeFormat(self.getTime())
                            excelUtil.insert_indicator_data(indiId,indiName,excelCycle,orgStdCode,orgStdName,'2',value1,value2,createTime)

    def get_std_num(self,value,unit):
        if unit == '万元':
            return float(value) * 10000
        elif unit == '十万元':
            return float(value) * 100000
        elif unit == '百万元':
            return float(value) * 1000000
        elif unit == '千万元':
            return float(value) * 10000000
        elif unit == '亿元':
            return float(value) * 100000000
        else:
            return float(value)

    #获取首行
    def get_first_line(self,file_cn_name,sheetName, excelCycle, cfg,table):
        # 读首行，行数据写入org_info
        print("获取行首机构信息...")
        excelUtil = ExcelUtil()
        if cfg:
            # 获取列的起始位置和结束位置
            print(cfg)
            colStart = cfg[13]
            colEnd = table.ncols#cfg[14]
            print("%s 页列起始位置为：%s,%s,%s ~ %s " % (sheetName,cfg[9],cfg[10],colStart, colEnd))
            # 写入之前删除该sheet页机构记录
            excelUtil.delete_org_info(file_cn_name,sheetName,excelCycle)
            i = 0
            for value in table.row_values(0, start_colx=colStart, end_colx=colEnd):
                print(value)
                # 插入新的机构记录
                excelUtil.insert_org_info(file_cn_name,sheetName,value,i+1,0,colStart+i,excelCycle)
                i += 1

    def get_std_org_info(self,srcOrgName):
        # 如果名称简化之后 只剩 地市名称，则直接匹配
        #tmp1 =
        excelUtil = ExcelUtil()
        return excelUtil.get_std_org_info(srcOrgName)


    #获取首列
    def get_first_col(self, file_cn_name, sheetName,excelCycle, cfg,table,):
        print("获取列首指标名称...")
        excelUtil = ExcelUtil()
        year = excelCycle[0:4]
        if cfg:
            #获取行的起始位置和结束位置
            print(cfg)
            lineStart = cfg[11]
            lineEnd = cfg[12]
            print("%s 页行起始位置为：%s,%s,%s ~ %s " % (sheetName, cfg[7], cfg[8], lineStart, lineEnd))
            # 写入之前删除该sheet页机构记录
            # 存在一个sheet页分多次读取的情况，所以删除需要按条件删除，存在之前已经删过一次的情况
            excelUtil.delete_indi_node(file_cn_name, sheetName, excelCycle,lineStart,lineEnd)
            i = 0
            for value in table.col_values(0, start_rowx=lineStart, end_rowx=lineEnd):
                print(value)
                value = value.replace('其中：','').replace('其中:','').strip()
                # 指标特例额处理，用于处理财务数据_欠费_速递账中因无编号导致指标名称重复而缺失问题
                # 给E标准年初欠费和年末欠费 加上序号
                valueTmp = value[0:2]
                if file_cn_name == '财务数据_欠费_速递账' and year == '2019':
                    if i == 55 :
                        if valueTmp != '1.':
                            value = '1.' + value
                    if i == 58 :
                        if valueTmp != '2.':
                            value = '2.' + value
                # 用于处理财务数据_收入预算_企业_寄递事业部中利润忘加序号而导致层级错乱问题
                if file_cn_name == '财务数据_收入预算_企业_寄递事业部' and year == '2019':
                    if i ==46:
                        if valueTmp != '二、':
                            value = '二、' + value
                    if i ==47:
                        if valueTmp != '三、':
                            value = '三、' + value
                # 插入新的指标记录
                try:
                    create_tm = self.timeFormat(self.getTime())
                    excelUtil.insert_indi_node(file_cn_name, sheetName, value, i + 1, lineStart + i, 0, excelCycle,create_tm)
                except:
                    print("插入indi_node出错！！！")
                    print("记录信息：file_cn_name：" + file_cn_name)
                    print("记录信息：sheetName：" + sheetName)
                    print("记录信息：excelCellSeq：" + lineStart + i)
                i += 1

    # 读取客户欠费详细数据
    def read_cust_owe_dtl(self,file_cn_name, sheetName, cfg, table,excelCycle):
        print("开始解析财务数据_欠费详细_企业...")
        excelUtil = ExcelUtil()
        if cfg:
            # 获取行的起始位置和结束位置
            print(cfg)
            lineStart = cfg[11]
            lineEnd = cfg[12]
            colStart = cfg[13]
            colEnd = cfg[14]
            print("%s 页行起始位置为：%s,%s,%s ~ %s " % (sheetName, cfg[7], cfg[8], lineStart, lineEnd))
            print("%s 页列起始位置为：%s,%s,%s ~ %s " % (sheetName, cfg[9], cfg[10], colStart, colEnd))
            # 写入之前删除该sheet页指标记录
            excelUtil.delete_cust_owe_dtl(excelCycle)
            # 有效行数和列数
            nrows = table.nrows
            print(nrows)
            ncols = table.ncols
            print(ncols)
            # 合并行
            #mergedLineStart = table.merged_cells[0]
            #mergedLineEnd = table.merged_cells[1]

            #print("合并行数为：%s ~ %s" %(mergedLineStart,mergedLineEnd))
            merged_cells = {}
            #merged_cells = table.merged_cells

            #有效行数为   nrows-2  开始行数为第7行，因为是从0开始，所以开始行数为6
            for i in range(lineStart,nrows-2):
                row_data = table.row_values(i)  # 按行获取excel的值
                # 0. empty（空的）,1 string（text）, 2 number, 3 date, 4 boolean, 5 error， 6 blank（空白表格）
                # 为了避免文本里面的数据出现带.0的情况，需要先转换为整型之后转换为字符串
                if table.cell(i,10).ctype == 2:
                    row_data[10] = str(int(row_data[10]))

                if row_data[4]:
                    row_data[4] = str(int(row_data[4]))
                try :
                    excelUtil.insert_cust_owe_dtl(row_data[0],row_data[1],row_data[2],row_data[3],row_data[4],
                                                  row_data[5],row_data[6],row_data[7],row_data[8],row_data[9],
                                                  row_data[10],row_data[11].strip(),row_data[12],row_data[13],row_data[14],
                                                  row_data[15],row_data[16],row_data[17],row_data[18],row_data[19],
                                                  row_data[20],excelCycle)
                except :
                    print('插入insert_cust_owe_dtl出错！！')
                    print("记录信息：当前行数：" + str(i))
                    print("记录信息：SRC_SYS_NM：" + row_data[0])
                    print("记录信息：ACCT_AGENCY_NM：" + row_data[1])
                    print("记录信息：DEPT_CD：" + row_data[2])
                    print("记录信息：DEPT_NM：" + row_data[3])
                    print("记录信息：WBST_CD：" + row_data[4])
                    print("记录信息：CUST_CD：" + row_data[11])


    # 获取unix时间戳
    def getTime(self):
        return round(time.time())

    # 时间格式化
    def timeFormat(self,timestamp):
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(timestamp))


#主程序入口
if __name__ == '__main__':
    ex = Excel()
    dirName = sys.argv[1]
    excelCycle = sys.argv[2]
    ex.read_excel(dirName,excelCycle)
    #ex.read_excel('/Volumes/Work/亚信/T07财务部数据/')
