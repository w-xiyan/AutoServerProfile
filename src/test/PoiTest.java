package test;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
 


import java.util.Map;






import oracle.net.aso.d;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import dao.BaseDaoDemo;
import util.ExportExcelUtils;
import util.GetTime;
import util.TimeUtil;
 
public class PoiTest {
    /**
     * @param	
     * @param
     * @param
     * @return
     * @author XIYAN
     */
	@SuppressWarnings("unchecked")
	public static void main(String[] args) {
			String schedule = "";
			String register = "";
			//��ȡδ��14��ʱ�䣬�������������
			ArrayList<String> list = TimeUtil.getTime(14);
			
			
			int i = 1 ;
			//�������飬��ȡʱ��
			for (String ls : list) {
				if (i == list.size()) {
					//ƴ�Ӻ�Դ�ַ���
					schedule = schedule +" SUM(CASE WHEN CLINIC_DATE = '" +ls +"' THEN RESOURCE_NUM END)";
					//ƴ���Ű��ַ���
					register = register +" SUM(CASE WHEN CLINIC_DATE = '" +ls +"' THEN RESOURCE_NUM END)";
				}else{
					schedule = schedule  +" SUM(CASE WHEN CLINIC_DATE = '" +ls +"' THEN RESOURCE_NUM END),";
					register = register  +" SUM(CASE WHEN CLINIC_DATE = '" +ls +"' THEN RESOURCE_NUM END),";
					i++;
				}
			}
			
			//1.��������Դ��ϸ
			String sql1 = "select  *��from (SELECT ROUTE ,ORG_NAME,"
						+	schedule
						+" FROM ("
						+" SELECT ROUTE ,ORG_NAME,CLINIC_DATE,RESOURCE_NUM"
						+"  FROM (SELECT T.ROUTE,"
						+" T.ORG_CODE,"
						+" T.CLINIC_DATE,"
						+" COUNT(0) RESOURCE_NUM"
						+" FROM RSM_REGISTER_INFO@DQCDBLINK T"
						+" GROUP BY T.ROUTE, T.ORG_CODE, T.CLINIC_DATE"
						+" ORDER BY T.CLINIC_DATE, T.ROUTE) A,"
						+" (SELECT D.ORG_CODE, D.ORG_NAME, D.ORG_ALIAS"
						+" FROM MDM_MD.MDM_ORGAN@DQCDBLINK D) B"
						+" WHERE A.ORG_CODE = B.ORG_CODE"
						+" ORDER BY A.CLINIC_DATE, A.ROUTE DESC�� C "
						+" GROUP BY ROUTE ,ORG_NAME"
						+" ORDER BY ROUTE ,ORG_NAME)";
			System.out.println(sql1);
			//2 .�������Ű���ϸ
			String sql2 = "select  *�� from (SELECT ROUTE ,ORG_NAME,DEPT_NAME,"
						+ register
						+" FROM ("	
						+" SELECT A.ROUTE,B.ORG_CODE, B.ORG_NAME,A.CLINIC_DATE,A.DEPT_NAME,A.SOURCE_SCHEDUL"	
						+" FROM (SELECT T.ROUTE,"	
						+" T.CLINIC_DATE,"	
						+" T.ORG_CODE,"	
						+" T.DEPT_CODE,"	
						+" T.DEPT_NAME,"	
						+" SUM(T.REG_TOTAL_NUM) SOURCE_SCHEDUL"	
						+" FROM RSM_SCHEDUL_INFO@DQCDBLINK T"	
						+" Where T.schedul_status<>'1'"	
						+" GROUP BY T.ROUTE,"	
						+" T.CLINIC_DATE,T.ORG_CODE,T.DEPT_CODE,T.DEPT_NAME"	
						+" ORDER BY T.ROUTE, T.CLINIC_DATE ASC) A,"	
						+" (SELECT D.ORG_CODE, D.ORG_NAME, D.ORG_ALIAS FROM MDM_MD.MDM_ORGAN@DQCDBLINK D) B"	
						+" WHERE A.ORG_CODE = B.ORG_CODE"	
						+" ORDER BY A.CLINIC_DATE,A.ROUTE DESC�� C"	
						+" GROUP BY ROUTE ,ORG_NAME,DEPT_NAME"	
						+" ORDER BY ROUTE ,ORG_NAME,DEPT_NAME)";

			
			//��ȡ�����
			List<Map<String,Object>> datalist1 = BaseDaoDemo.findListBySql(sql1);
			List<Map<String,Object>> datalist2 = BaseDaoDemo.findListBySql(sql2);
			
			//����sheet��ͷ
			ArrayList<String> list1 = new ArrayList<String>();
			list1 = list ;
			list1.add(0, "����ID");
			list1.add(1, "��������");
			System.out.println(list1);
			ArrayList<String> list2 = new ArrayList<String>();
			list2 = list ;
			list2.add(0, "����ID");
			list2.add(1, "��������");
			list2.add(2, "��������");
			list2.remove(3);
			list2.remove(3);
			System.out.println(list2);
			
			//String[] header1 = { "����ID", "��������", "ʧ�ܴ�������)","ƽ����ʱ���룩" ,"ִ���ܴ������Σ�","�ɹ���" };
			
			//String[] header2 = { "��������","��������","ʧ�ܴ�������)","ƽ����ʱ����)","ԭ��" };
			
			ExportExcelUtils eeu1 = new ExportExcelUtils();
			HSSFWorkbook workbook1 = new HSSFWorkbook();
			ExportExcelUtils eeu2 = new ExportExcelUtils();
			HSSFWorkbook workbook2 = new HSSFWorkbook();

			try {
				eeu1.exportExcel(workbook1, 0, "��Դ", list1, datalist1);
				eeu2.exportExcel(workbook2, 0, "�Ű�", list2, datalist2);
				
			} catch (Exception e) {
				e.printStackTrace();
			}
	}
}
