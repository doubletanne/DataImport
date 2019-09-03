
#文件配置表，excel_cfg，数据为手工配置
class ExcelCfg():
    def __init__(self, excelName, excelCnName, sheetNames, year, excelFreq, excelCycle):
        self.excelName = excelName
        self.excelCnName = excelCnName
        self.sheetNames = sheetNames
        self.year = year
        self.excelFreq = excelFreq
        self.excelCycle = excelCycle

#文件sheet名称配置表，file_sheet_cfg，数据为手工配置，一个文件有多个sheet页
class FileSheetCfg():
    def __init__(self, excelName, sheetName, sheetDataFreq, acctBookType, acctType, sheetSeqPrefix, sheetUnit,
                 rowStartNum, rowEndNum, colStartNum, colEndNum,year):
        self.excelName = excelName
        self.sheetName = sheetName
        self.sheetDataFreq = sheetDataFreq
        self.acctBookType = acctBookType
        self.acctType = acctType
        self.sheetSeqPrefix = sheetSeqPrefix
        self.sheetUnit = sheetUnit
        self.rowStartNum = rowStartNum
        self.rowEndNum = rowEndNum
        self.colStartNum = colStartNum
        self.colEndNum = colEndNum
        self.year = year

#文件读取历史表，file_read_log，数据为写入
class ExcelReadLog():
    def __init__(self,file_nm,sheet_nm,file_cn_nm,file_cycle,file_read_tm):
        self.file_nm = file_nm
        self.sheet_nm = sheet_nm
        self.file_cn_nm = file_cn_nm
        self.file_cycle = file_cycle
        self.file_read_tm = file_read_tm

#指标树，indi_tree，数据为写入
# class IndiTree():
#     def __init__(self,file_nm=None,sheet_nm=None,indi_id=None,indi_nm=None,indi_lvl=None,indi_seq=None,data_cycle=None,excel_cell_seq=None,
#                  indi_lvl1_nm=None,indi_lvl2_nm=None,indi_lvl3_nm=None,indi_lvl4_nm=None,indi_lvl5_nm=None,indi_lvl6_nm=None,
#                  indi_lvl7_nm=None,indi_lvl8_nm=None):
#         self.file_nm = file_nm
#         self.sheet_nm = sheet_nm
#         self.indi_id = indi_id
#         self.indi_nm = indi_nm
#         self.indi_lvl = indi_lvl
#         self.indi_seq = indi_seq
#         self.data_cycle = data_cycle
#         self.excel_cell_seq = excel_cell_seq
#         self.indi_lvl1_nm = indi_lvl1_nm
#         self.indi_lvl2_nm = indi_lvl2_nm
#         self.indi_lvl3_nm = indi_lvl3_nm
#         self.indi_lvl4_nm = indi_lvl4_nm
#         self.indi_lvl5_nm = indi_lvl5_nm
#         self.indi_lvl6_nm = indi_lvl6_nm
#         self.indi_lvl7_nm = indi_lvl7_nm
#         self.indi_lvl8_nm = indi_lvl8_nm
#指标名称记录
class IndiNode():
    def __init__(self,file_nm,sheet_nm,indi_nm,indi_seq,excel_line_num,excel_col_num):
        self.file_nm = file_nm
        self.sheet_nm = sheet_nm
        self.indi_nm = indi_nm
        self.indi_seq = indi_seq
        self.excel_line_num = excel_line_num
        self.excel_col_num = excel_col_num


#机构记录，org_info，数据为写入
class OrgInfo():
    def __init__(self,file_nm,sheet_nm,org_nm,org_seq,excel_line_num,excel_col_num):
        self.file_nm = file_nm
        self.sheet_nm = sheet_nm
        self.org_nm = org_nm
        self.org_seq = org_seq
        self.excel_line_num = excel_line_num
        self.excel_col_num = excel_col_num

#标准机构信息，std_org_info，数据为手工配置
class StdOrgInfo():
    def __init__(self,org_id,par_org_id,org_nm,group_flag):
        self.org_id = org_id
        self.par_org_id = par_org_id
        self.org_nm = org_nm
        self.group_flag = group_flag

#邮政客户欠费详细表，STAGE_POST_CUST_OWE_DTL_YYYYMM
class CustOweDtl():
    def __init__(self,SRC_SYS_NM,ACCT_AGENCY_NM,DEPT_CD,DEPT_NM,WBST_CD,WBST_NM,SPEC_CD,SPEC_NM,PRD_CD,PRD_NM,RSPS_PSN_CD,CUST_CD,CUST_NM,RECPT_NM,OWE_IN_1_M,OWE_1_3_M,OWE_3_6_M,OWE_6_12_M,OWE_1_3_Y,OWE_MT_3_Y,OWE_TOTAL,statis_month):
        self.SRC_SYS_NM = SRC_SYS_NM
        self.ACCT_AGENCY_NM = ACCT_AGENCY_NM
        self.DEPT_CD = DEPT_CD
        self.DEPT_NM = DEPT_NM
        self.WBST_CD = WBST_CD
        self.WBST_NM = WBST_NM
        self.SPEC_CD = SPEC_CD
        self.SPEC_NM = SPEC_NM
        self.PRD_CD = PRD_CD
        self.PRD_NM = PRD_NM
        self.RSPS_PSN_CD = RSPS_PSN_CD
        self.CUST_CD = CUST_CD
        self.CUST_NM = CUST_NM
        self.RECPT_NM = RECPT_NM
        self.OWE_IN_1_M = OWE_IN_1_M
        self.OWE_1_3_M = OWE_1_3_M
        self.OWE_3_6_M = OWE_3_6_M
        self.OWE_6_12_M = OWE_6_12_M
        self.OWE_1_3_Y = OWE_1_3_Y
        self.OWE_MT_3_Y = OWE_MT_3_Y
        self.OWE_TOTAL = OWE_TOTAL
        self.statis_month = statis_month
#指标数据
class IndiData():
    def __init__(self,seq,kpi_id,kpi_nm,cycle_id,org_id,org_nm,org_lvl,value1,value2):
        self.seq = seq
        self.kpi_id = kpi_id
        self.kpi_nm = kpi_nm
        self.cycle_id = cycle_id
        self.org_id = org_id
        self.org_nm = org_nm
        self.org_lvl = org_lvl
        self.value1 = value1
        self.value2 = value2

