from pgdb import *

class ExcelUtil():
    def __init__(self):
         self.dbMgr = pgDbMgr()

    """获取excel的配置，返回list"""
    def get_excel_cfg(self,fileCnName,year):
        return self.dbMgr.fetchall(table='excel_cfg',field=' * ',where=" file_nm='"+fileCnName+"' and year='"+year+"'")


    def is_excel_exist(self,fileCnName,year):
        if self.dbMgr.fetchone(table='excel_cfg',field=' * ',where=" file_nm='"+fileCnName+"' and year='"+year+"'"):
            return True
        else:
            return False

    """获取excel sheet页的配置，返回list"""
    def get_excel_sheet_cfg(self,fileCnName,sheetName,year):
        return self.dbMgr.fetchall(table='excel_sheet_cfg',field=' * ',where=" file_nm='"+fileCnName+ "' and sheet_nm='" + sheetName +"' and  year='"+year+"'")

    def is_sheet_exist(self,fileCnName,sheetName,year):
        if self.dbMgr.fetchall(table='excel_sheet_cfg',field=' * ',where=" file_nm='"+fileCnName+ "' and sheet_nm='" + sheetName +"' and  year='"+year+"'"):
            return True
        else:
            return False

    """获取sheet页的开始列号和结束列号"""
    def get_sheet_col_cell(self, fileCnName, sheetName,year):
        return self.dbMgr.fetchall(table='excel_sheet_cfg',field=' col_start_num,col_end_num ',where=" file_nm='"+fileCnName+ "' and sheet_nm='" + sheetName +"' and  year='"+year+"'")

    """获取sheet页的开始行号和结束行号"""
    def get_sheet_line_cell(self, fileCnName, sheetName,year):
        return self.dbMgr.fetchall(table='excel_sheet_cfg',field=' row_start_num,row_end_num ',where=" file_nm='"+fileCnName+ "' and sheet_nm='" + sheetName +"' and  year='"+year+"'")

    """获取指标前缀"""
    """
    fileCnName 文件中文名，不带序号和账期
    sheetName 文件sheet名称
    excelCell 单元格号
    fileFreq  文件频率，年还是月
    """
    def get_indi_prefix(self, fileCnName, sheetName,indiLineCell,year):
        if int(indiLineCell) > 0:
            return self.dbMgr.fetchone(table='excel_sheet_cfg',field=' sheet_seq_prefix ',
                                   where=" file_nm='"+fileCnName+ "' and sheet_nm='" + sheetName +"' and  year='"+year+"' and  row_start_num <= " + str(indiLineCell) + "::int and row_end_num >= " + str(indiLineCell) +"::int" )
        else:
            return ''

    """判断指标信息是否已经存在"""
    def get_indi_id(self,fileCnName,sheetName,indiName,indiLvl1Name,indiLvl2Name,indiLvl3Name,indiLvl4Name,indiLvl5Name,indiLvl6Name,indiLvl7Name,indiLvl8Name):
        return self.dbMgr.fetchone(table='indi_tree',field= ' indi_id ',
                               where="file_nm='" + fileCnName + "' and sheet_nm='" + sheetName + "' and indi_nm='" + indiName +
                                     "' and indi_lvl1_nm='" + indiLvl1Name + "' and coalesce(indi_lvl2_nm,'')='" + indiLvl2Name + "' and coalesce(indi_lvl3_nm,'')='" + indiLvl3Name + "' and coalesce(indi_lvl4_nm,'')='" + indiLvl4Name +
                                     "' and coalesce(indi_lvl5_nm,'')='" + indiLvl5Name + "' and coalesce(indi_lvl6_nm,'')='" + indiLvl6Name + "' and coalesce(indi_lvl7_nm,'')='" + indiLvl7Name + "' and coalesce(indi_lvl8_nm,'')='" + indiLvl8Name + "'")

    """插入indi_tree"""
    def insert_indi_tree(self,fileCnName,sheetName,indiId,indiName,indiLvl,indiLvl1Name,
                         indiLvl2Name,indiLvl3Name,indiLvl4Name,indiLvl5Name,indiLvl6Name,indiLvl7Name,indiLvl8Name,create_tm
                         ):
        self.dbMgr.insert(table='indi_tree', file_nm=fileCnName,sheet_nm=sheetName,indi_id=indiId,indi_nm=indiName,indi_lvl=indiLvl
                          ,indi_lvl1_nm=indiLvl1Name,indi_lvl2_nm=indiLvl2Name,indi_lvl3_nm=indiLvl3Name,indi_lvl4_nm=indiLvl4Name,
                          indi_lvl5_nm=indiLvl5Name,indi_lvl6_nm=indiLvl6Name,indi_lvl7_nm=indiLvl7Name,indi_lvl8_nm=indiLvl8Name,create_tm=create_tm
                          )

    """返回序列最大值"""
    def get_indi_seq_max(self,indiSeqPrefix):
        if indiSeqPrefix:
            seq = self.dbMgr.fetchone(table='indi_tree',field=' max(indi_id) ',
                                       where=" indi_id like '" + indiSeqPrefix + "%'")[0].__str__()
            if seq != 'None' and seq != None:
                print('seq:' + seq)
                seqMax = seq[-5:]
                print('seqMax:' + seqMax)
                return seqMax
            else:
                return '0'

    """插入文件读取日志表中"""
    def insert_file_read_log(self,fileName,sheetName,fileCnName,fileCycle,fileReadTime):
        #
        self.dbMgr.insert(table='file_read_log', file_nm=fileName, sheet_nm=sheetName, file_cn_nm=fileCnName,
               file_cycle=fileCycle, file_read_tm=fileReadTime)

    """删除sheet页的机构信息"""
    def delete_org_info(self,fileCnName,sheetName,excelCycle):
        rows = self.dbMgr.delete(table='org_info',where="file_nm='" + fileCnName + "' and sheet_nm='" + sheetName + "'  and cycle_id='"+excelCycle+"'")
        print("删除了 %d " % rows + "行。")

    """插入sheet页的机构信息"""
    def insert_org_info(self, fileCnName, sheetName,orgName,orgSeq,excelLineNum, excelColNum, excelCycle):
        self.dbMgr.insert(table='org_info', file_nm=fileCnName,sheet_nm=sheetName,org_nm=orgName,org_seq=orgSeq,excel_line_num=excelLineNum,excel_col_num=excelColNum,cycle_id=excelCycle)

    """获取机构名称信息"""
    def get_org_names(self,fileCnName,sheetName,excelCycle):
        return self.dbMgr.fetchall(table='org_info',field=' org_nm,excel_col_num ',where=" file_nm='"+fileCnName+ "' and sheet_nm='" + sheetName +"' and  cycle_id='"+excelCycle+"'")

    """获取标准机构名称代码信息"""
    def get_std_org_info(self,srcOrgName):
        return self.dbMgr.fetchone(table='org_info_map', field=' std_org_id,std_org_nm ',
                                   where=" src_org_nm='" + srcOrgName + "'")

    """删除sheet页的指标名称信息"""

    def delete_indi_node(self, fileCnName, sheetName,cycle_id,lineStart,lineEnd):
        rows = self.dbMgr.delete(table='indi_node',
                       where="file_nm='" + fileCnName + "' and sheet_nm='" + sheetName + "'  and cycle_id='"+cycle_id+"' and excel_line_num::int >= " + str(lineStart) + " and excel_line_num::int < " + str(lineEnd)  + " ")
        print("删除了 %d " % int(rows) + "行。")

    """插入sheet页的指标名称信息"""

    def insert_indi_node(self, fileCnName, sheetName, indiName, indiSeq, excelLineNum, excelColNum, cycleId,createTm):
        self.dbMgr.insert(table='indi_node', file_nm=fileCnName, sheet_nm=sheetName, indi_nm=indiName, indi_seq=indiSeq,
                       excel_line_num=excelLineNum, excel_col_num=excelColNum,cycle_id=cycleId,create_tm=createTm)

    """获取指标名称信息"""
    def get_indi_node(self, fileCnName, sheetName,cycleId):
        return self.dbMgr.fetchall(table='indi_node', field=' indi_nm,indi_seq,excel_line_num,node_id ',
                       where=" file_nm='" + fileCnName + "' and sheet_nm='" + sheetName + "'  and cycle_id='"+cycleId+"'", order=" excel_line_num::int")

    """获取指标的节点表的指标ID信息"""
    def get_indi_node_indi_id(self,fileCnName, sheetName, cycleId, rowStart, rowEnd):
        return self.dbMgr.fetchall(table='indi_node', field=' indi_id,indi_nm_noseq,excel_line_num ',
                                   where=" file_nm='" + fileCnName + "' and sheet_nm='" + sheetName + "'  and cycle_id='" + cycleId + "'  and excel_line_num::int >= " + str(rowStart)+ " and excel_line_num::int < "+ str(rowEnd) +" ",
                                   order=" excel_line_num::int")

    def update_indi_node(self, nodeId,indiId,indiNm):
        return self.dbMgr.update(table='indi_node', indi_id=indiId,indi_nm_noseq=indiNm,
                       where=" node_id=" + str(nodeId) + " ")


    """删除指定账期客户欠费详细信息"""
    def delete_cust_owe_dtl(self, excelCycle):
        rows = self.dbMgr.delete(table='STAGE_POST_CUST_OWE_DTL_YYYYMM',
                                 where="statis_month='" + excelCycle + "'")
        print("删除了 %d " % rows + "行。")

    """插入指定账期客户欠费详细信息"""
    def insert_cust_owe_dtl(self,SRC_SYS_NM,ACCT_AGENCY_NM,DEPT_CD,DEPT_NM,WBST_CD,WBST_NM,SPEC_CD,SPEC_NM,PRD_CD,PRD_NM,RSPS_PSN_CD,
                            CUST_CD,CUST_NM,RECPT_NM,OWE_IN_1_M,OWE_1_3_M,OWE_3_6_M,OWE_6_12_M,OWE_1_3_Y,OWE_MT_3_Y,OWE_TOTAL,statis_month):
        self.dbMgr.insert(table='STAGE_POST_CUST_OWE_DTL_YYYYMM',SRC_SYS_NM=SRC_SYS_NM,ACCT_AGENCY_NM=ACCT_AGENCY_NM,DEPT_CD=DEPT_CD,DEPT_NM=DEPT_NM,
                          WBST_CD=WBST_CD,WBST_NM=WBST_NM,SPEC_CD=SPEC_CD,SPEC_NM=SPEC_NM,PRD_CD=PRD_CD,PRD_NM=PRD_NM,RSPS_PSN_CD=RSPS_PSN_CD,
                            CUST_CD=CUST_CD,CUST_NM=CUST_NM,RECPT_NM=RECPT_NM,OWE_IN_1_M=OWE_IN_1_M,OWE_1_3_M=OWE_1_3_M,OWE_3_6_M=OWE_3_6_M,
                          OWE_6_12_M=OWE_6_12_M,OWE_1_3_Y=OWE_1_3_Y,OWE_MT_3_Y=OWE_MT_3_Y,OWE_TOTAL=OWE_TOTAL,statis_month=statis_month)

    """删除指标数据表中已经存在指标数据"""
    def delete_indicator_data_exists(self,file_cn_name,sheetName,excelCycle,rowStart,rowEnd):
        rows = self.dbMgr.delete(table='indicator_data',where=" kpi_id in (select indi_id from indi_node where file_nm='" + file_cn_name +"' and sheet_nm='"+sheetName+"' and excel_line_num::int >= " + str(rowStart)+ " and excel_line_num::int < "+ str(rowEnd) +" and cycle_id='" + excelCycle + "') and cycle_id='" + excelCycle + "' ")
        print("删除了 %d " % rows + "行。")

    """插入指标数据"""

    def insert_indicator_data(self,kpi_id,kpi_nm,cycle_id,org_id,org_nm,org_lvl,value1,value2,create_tm):
        self.dbMgr.insert(table='indicator_data',kpi_id=kpi_id,kpi_nm=kpi_nm,cycle_id=cycle_id,org_id=org_id,org_nm=org_nm,org_lvl=org_lvl,value1=value1,value2=value2,create_tm=create_tm)

