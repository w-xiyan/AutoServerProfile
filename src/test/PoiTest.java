package test;

import java.util.Calendar;
import java.util.Date;
import java.util.List;
 


import java.util.Map;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import dao.BaseDaoDemo;
import util.ExportExcelUtils;
import util.GetTime;
 
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
			Calendar cal=Calendar.getInstance();
			cal.setTime(new Date()); 
			int week=cal.get(Calendar.DAY_OF_WEEK)-1;
			String datetime1;
			String datetime2;
			String datetime3;
			String datetime4;
			String datetime5;
			String datetime6;
			if (week ==1) {
				//��ȡ���������� ��ʽΪ	yyyy-MM-dd HH:mm:ss
				 datetime1 = GetTime.getBeforeThreeTime();
				//�������� ��ʽΪ yyyy-MM-dd HH:mm:ss
				 datetime2 = GetTime.getnowEndTime();
				//��ȡ���������� ��ʽΪ	yyyyMMdd
				 datetime3 = GetTime.getBeforeThreeToTime();
				//�������� ��ʽΪ	 yyyyMMdd
				 datetime4 = GetTime.getTodayTime();
				//�������� ��ʽΪ	 yyyyMMdd
				 datetime6 = GetTime.getTomorrowTime();
				//����һ������ 	��ʽΪyyyyMMdd
				 datetime5 = GetTime.getWeekDayTime();
			}else {
				//��������
				 datetime1 = GetTime.getStartTime();
				//��������
				 datetime2 = GetTime.getnowEndTime();
				//��������
				 datetime3 = GetTime.getYesterdayTime();
				//��������
				 datetime4 = GetTime.getTodayTime();
				//����һ������
				 datetime5 = GetTime.getWeekDayTime();
				//�������� ��ʽΪ	 yyyyMMdd
				 datetime6 = GetTime.getTomorrowTime();
			}
			
			//1.Ѳ��HSB����sql-1-�ſ�
			String sql1 = "";

			//2 .Ѳ��HSB����sql-1-ʧ��ͳ�ƣ���ԭ��
			String sql2 = "SELECT N.NODE_NAME,"
			  		+  "S.SEIN_CNAME,"
			  		+  "T.FAILED_COUNT,"
			  		+  "T.AVG_TIME,"
			  		+	"T.ERR_NAME"
			  		+  " FROM hsb_platform.HSB_REG_NODEINFO@DQCDBLINK N,"
			  		+  "hsb_platform.SERVICE_INFO@DQCDBLINK S,"
			  		+  "(SELECT TAR_SERIVCE_ID,TAR_NODE_ID,ERR_NAME,(1) AS FAILED_COUNT,ROUND (AVG(TIME_COST)/1000, 1) AS AVG_TIME"
			  		+  " FROM hsb_platform.log_services@DQCDBLINK WHERE RESULT_CODE= - 1 "
			  		+  "AND (CALL_TIME BETWEEN TO_DATE ("
			  		+	"'"+datetime1+"'"
			  		+	",'YYYY-MM-DD HH24:MI:SS')"
			  		+  "AND TO_DATE ("
			  		+	"'"+datetime2+"'"
			  		+	",'YYYY-MM-DD HH24:MI:SS'))"
			  		+	"AND TAR_NODE_ID IN ("
					+  "'7a40f95c-bec6-4163-86ed-08e00cdfb073',"
			  		+  "'920f3d9d-86e3-4312-bf76-1cac906e4826',"
				    +  "'06f4b7bb-7325-4ccd-91a7-0a38ad1727ac',"
				    +  "'3788a2b8-3eb2-4c88-859d-c8df2dbc7410',"
				    +  "'c5ed05ca-2a3f-4c66-bdfc-facc0b5e2437',"
				    +  "'73b8fc7b-1fa5-466f-bcdf-dcf4ee6127f0',"
				    +  "'b0ef97ab-ec67-49e3-9867-e87e8a27512e',"
				    +  "'8bee4841-5ceb-426b-a19c-c6a77d8155c9',"
				    +  "'ec60389a-1aa4-4b83-8b75-6a3258644904',"
				    +  "'cfb5f485-6f90-4d34-946c-b24c1b8efc85',"
				    +  "'37719c69-e088-42c2-9ba2-6ba4debb1e2e',"
				    +  "'96704a80-ab3d-4f3f-a5ca-c98ba23bd81d',"
				    +  "'db7056bf-ef7f-4c6d-8fac-6e73adc2cb83',"
				    +  "'63e378b6-9bae-48d5-b310-ebeb8f045ba6',"
				    +  "'9551166b-6f64-493c-a59e-61fee86aa8f2',"
				    +  "'ada4f9d0-63b5-414c-b015-269d4013c621',"
				    +  "'6da5c2e1-ee20-4f74-a220-d1593b2c2f52')"
	    		+ " GROUP BY TAR_NODE_ID,TAR_SERIVCE_ID,ERR_NAME) T"
	    		+ " WHERE n.NODE_ID = T .TAR_NODE_ID AND s.SEIN_SERVICE_ID = T .TAR_SERIVCE_ID"
	    		+ " order by NODE_NAME, SEIN_CNAME";

			//�ɹ�
			String sql3 = "select * from ( SELECT "
			  		 +	"N.NODE_NAME,"
					 +	"N.SEIN_CNAME,"
			  		 +	"T.FAILED_COUNT,"
					 +	"T.AVG_TIME"
			  		 +	" FROM (select "
					 +	"n.node_id as NODE_ID,"
					 +	"n.node_name as NODE_NAME,"
					 +	"s.sein_service_id as SEIN_SERVICE_ID,"
					 +	"s.sein_cname as SEIN_CNAME"
					 +	" from "
					 +	"hsb_platform.HSB_REG_NODEINFO@DQCDBLINK n,"
					 +	"hsb_platform.SERVICE_INFO@DQCDBLINK s "
					 +	" where n.node_id = s.sein_node_id"
					 +	")N "
					 +	" left join("
					 +	" SELECT "
					 +	" TAR_SERIVCE_ID,"
					 +	"TAR_NODE_ID,"
					 +	"COUNT (1) AS FAILED_COUNT,"
					 +	" ROUND (AVG(TIME_COST)/1000, 1) AS AVG_TIME"
					 +	" FROM "
					 +	" hsb_platform.LOG_SERVICES@DQCDBLINK"
					 +	" WHERE "
					 +	"(CALL_TIME BETWEEN TO_DATE("
					 +	"'"+datetime1+"'"
					 +  ",'YYYY-MM-DD HH24:MI:SS')"
					 +	"AND TO_DATE("
					 +	"'"+datetime2+"'"
					 +  ",'YYYY-MM-DD HH24:MI:SS'))"
					 +	" GROUP BY "
					 +	"TAR_NODE_ID,"
					 +	"TAR_SERIVCE_ID)T"
					 +	" on "
					 +	"N.NODE_ID = T .TAR_NODE_ID"
					 +	" AND N.SEIN_SERVICE_ID = T.TAR_SERIVCE_ID"
					 +	" order by NODE_NAME, SEIN_CNAME)n where n.failed_count is not null";

			//4.Ѳ��HSB����sql-����������
			String sql4 = "SELECT N.NODE_NAME,N.SEIN_CNAME,T .FAILED_COUNT,T .AVG_TIME"
					+	" FROM (select n.node_id as NODE_ID,n.node_name as NODE_NAME,s.sein_service_id as SEIN_SERVICE_ID,"
					+	"s.sein_cname as SEIN_CNAME"
					+	" from hsb_platform.HSB_REG_NODEINFO@DQCDBLINK n,hsb_platform.SERVICE_INFO@DQCDBLINK s "
					+	" where n.node_id = s.sein_node_id"
					+	")N left join(SELECT TAR_SERIVCE_ID,TAR_NODE_ID,COUNT (1) AS FAILED_COUNT,ROUND (AVG(TIME_COST)/1000, 1) AS AVG_TIME"
					+	" FROM hsb_platform.LOG_SERVICES@DQCDBLINK"
					+	" WHERE(CALL_TIME BETWEEN TO_DATE ("
					+	"'"+datetime1+"'"
					+	" ,'YYYY-MM-DD HH24:MI:SS')AND TO_DATE ("
					+	"'"+datetime2+"'"
					+	" ,'YYYY-MM-DD HH24:MI:SS'))GROUP BY TAR_NODE_ID,TAR_SERIVCE_ID) T"
					+	" on N.NODE_ID = T .TAR_NODE_ID"
					+	" AND N.SEIN_SERVICE_ID = T .TAR_SERIVCE_ID "
					+	" order by NODE_NAME, SEIN_CNAME ";
			//5.�ɼ���������
			String sql5 = "select source_org_name, class_name, count(1) COUNTNUM "
					+	" FROM  dc_index.bs_basic_index@DQCDBLINK t"
					+	" where to_date(substr(t.create_date, 0, 8), 'yyyymmdd') >="
					+	" to_date("
					+	"'"+datetime3+"'"
					+	", 'yyyymmdd') and"
					+	"  to_date(substr(t.create_date, 0, 8), 'yyyymmdd') <"
					+	"  to_date("
					+	"'"+datetime4+"'"
					+	", 'yyyymmdd') "
					+	" group by source_org_name, class_name"
					+	" order  by  class_name";
			
			String sql6 = "";
			
			//7.��ѯ��־���
			String sql7 = "";
			if (week == 1) {
				sql7 = "select m.DOMAIN_NAME, t.class_name, count(*) COUNTNUM"
						+	" from DC_INDEX.INDEX_VERIFY_LOG@DQCDBLINK t, DC_INDEX.Z_MAP_DOMAINCODENAME@DQCDBLINK m"
						+	" where to_date(substr(t.create_date, 0, 8), 'yyyymmdd') >="
						+	" to_date("
						+	"'"+datetime3+"'"
						+	", 'yyyymmdd') and"
						+	"  to_date(substr(t.create_date, 0, 8), 'yyyymmdd') <"
						+	"  to_date("
						+	"'"+datetime4+"'"
						+	", 'yyyymmdd') "
						+	" AND t.DOMAIN_CODE = m.DOMAIN_CODE"
						+	" group by m.DOMAIN_NAME, t.class_name"
						+	"  order by m.DOMAIN_NAME";
			}else {
				sql7 = "select m.DOMAIN_NAME, t.class_name, count(*) COUNTNUM"
						+	" from DC_INDEX.INDEX_VERIFY_LOG@DQCDBLINK t, DC_INDEX.Z_MAP_DOMAINCODENAME@DQCDBLINK m"
						+	" where substr(t.create_date, 1, 8) = "
						+	"'"+datetime3+"'"
						+	" AND t.DOMAIN_CODE = m.DOMAIN_CODE"
						+	" group by m.DOMAIN_NAME, t.class_name"
						+	"  order by m.DOMAIN_NAME";
			}
			
			//1.0.0Ѳ��HSB����sql-1-�ſ�
			String sql11 = "SELECT n.NODE_NAME,"
			  		+  "s.SEIN_CNAME,"
			  		+  "T.ALL_COUNT,"
			  		+  "T.AVG_TIME" 
			  		+  " FROM " +"hsb_platform.HSB_REG_NODEINFO@DQCDBLINK n ,"
			  		+  "hsb_platform.SERVICE_INFO@DQCDBLINK s ,"
			  		+  "(SELECT s.tar_node_id as NODE_ID,s.tar_serivce_id AS SERIVCE_ID,count(*) as all_count,ROUND (AVG(TIME_COST)/1000, 1) AS AVG_TIME"
			  		+  " FROM "
			  		+	" hsb_platform.log_services@DQCDBLINK S WHERE s.tar_serivce_id in ("
			  		+	" select distinct(l.tar_serivce_id) from hsb_platform.log_services@DQCDBLINK l where l.result_code = -1"
			  		+  "AND (l.CALL_TIME BETWEEN TO_DATE ("
			  		+	"'"+datetime1+"'"
			  		+	",'YYYY-MM-DD HH24:MI:SS')"
			  		+  "AND TO_DATE ("
			  		+	"'"+datetime2+"'"
			  		+	",'YYYY-MM-DD HH24:MI:SS'))"
			  		+  "AND l.TAR_NODE_ID IN ("
					+  "'7a40f95c-bec6-4163-86ed-08e00cdfb073',"
			  		+  "'920f3d9d-86e3-4312-bf76-1cac906e4826',"
				    +  "'06f4b7bb-7325-4ccd-91a7-0a38ad1727ac',"
				    +  "'3788a2b8-3eb2-4c88-859d-c8df2dbc7410',"
				    +  "'c5ed05ca-2a3f-4c66-bdfc-facc0b5e2437',"
				    +  "'73b8fc7b-1fa5-466f-bcdf-dcf4ee6127f0',"
				    +  "'b0ef97ab-ec67-49e3-9867-e87e8a27512e',"
				    +  "'8bee4841-5ceb-426b-a19c-c6a77d8155c9',"
				    +  "'ec60389a-1aa4-4b83-8b75-6a3258644904',"
				    +  "'cfb5f485-6f90-4d34-946c-b24c1b8efc85',"
				    +  "'37719c69-e088-42c2-9ba2-6ba4debb1e2e',"
				    +  "'96704a80-ab3d-4f3f-a5ca-c98ba23bd81d',"
				    +  "'db7056bf-ef7f-4c6d-8fac-6e73adc2cb83',"
				    +  "'63e378b6-9bae-48d5-b310-ebeb8f045ba6',"
				    +  "'9551166b-6f64-493c-a59e-61fee86aa8f2',"
				    +  "'ada4f9d0-63b5-414c-b015-269d4013c621',"
				    +  "'6da5c2e1-ee20-4f74-a220-d1593b2c2f52')"
				    +	")and s.CALL_TIME BETWEEN TO_DATE ("
				    +	"'"+datetime1+"'"
        			+	",'YYYY-MM-DD HH24:MI:SS')"
        			+	" AND TO_DATE ("
        			+	"'"+datetime2+"'"
        			+	",'YYYY-MM-DD HH24:MI:SS')"
        			+	"group by s.tar_node_id,s.tar_serivce_id) T"
        			+	" WHERE "+"n.NODE_ID = T.NODE_ID AND s.SEIN_SERVICE_ID = T.SERIVCE_ID"
        			+ 	" order by "+"NODE_NAME, SEIN_CNAME";
			
			//1.0.1Ѳ��HSB����sql-1 - �ܴ���
			String sql12 = "SELECT n.NODE_NAME,"
			  		+  "s.SEIN_CNAME,"
			  		+  "T.FAILED_COUNT,"
			  		+  "T.AVG_TIME" 
			  		+  " FROM " +"hsb_platform.HSB_REG_NODEINFO@DQCDBLINK n ,"
			  		+  "hsb_platform.SERVICE_INFO@DQCDBLINK s ,"
			  		+  "(SELECT TAR_SERIVCE_ID,TAR_NODE_ID,COUNT (1) AS FAILED_COUNT,ROUND (AVG(TIME_COST)/1000, 1) AS AVG_TIME"
			  		+  " FROM "
			  		+	" hsb_platform.log_services@DQCDBLINK WHERE RESULT_CODE= - 1"
			  		+  "AND (CALL_TIME BETWEEN TO_DATE ("
			  		+	"'"+datetime1+"'"
			  		+	",'YYYY-MM-DD HH24:MI:SS')"
			  		+  "AND TO_DATE ("
			  		+	"'"+datetime2+"'"
			  		+	",'YYYY-MM-DD HH24:MI:SS'))"
			  		+  "AND TAR_NODE_ID IN ("
					+  "'7a40f95c-bec6-4163-86ed-08e00cdfb073',"
			  		+  "'920f3d9d-86e3-4312-bf76-1cac906e4826',"
				    +  "'06f4b7bb-7325-4ccd-91a7-0a38ad1727ac',"
				    +  "'3788a2b8-3eb2-4c88-859d-c8df2dbc7410',"
				    +  "'c5ed05ca-2a3f-4c66-bdfc-facc0b5e2437',"
				    +  "'73b8fc7b-1fa5-466f-bcdf-dcf4ee6127f0',"
				    +  "'b0ef97ab-ec67-49e3-9867-e87e8a27512e',"
				    +  "'8bee4841-5ceb-426b-a19c-c6a77d8155c9',"
				    +  "'ec60389a-1aa4-4b83-8b75-6a3258644904',"
				    +  "'cfb5f485-6f90-4d34-946c-b24c1b8efc85',"
				    +  "'37719c69-e088-42c2-9ba2-6ba4debb1e2e',"
				    +  "'96704a80-ab3d-4f3f-a5ca-c98ba23bd81d',"
				    +  "'db7056bf-ef7f-4c6d-8fac-6e73adc2cb83',"
				    +  "'63e378b6-9bae-48d5-b310-ebeb8f045ba6',"
				    +  "'9551166b-6f64-493c-a59e-61fee86aa8f2',"
				    +  "'ada4f9d0-63b5-414c-b015-269d4013c621',"
				    +  "'6da5c2e1-ee20-4f74-a220-d1593b2c2f52')"
				    +	"GROUP BY TAR_NODE_ID,TAR_SERIVCE_ID) T"
				    +	" WHERE "+"n.NODE_ID = T .TAR_NODE_ID AND s.SEIN_SERVICE_ID = T .TAR_SERIVCE_ID"
				    +	" order by "+"NODE_NAME, SEIN_CNAME";
				sql1=	" SELECT t1.NODE_NAME,t1.SEIN_CNAME,t1.FAILED_COUNT,t1.AVG_TIME,t2.all_count FROM ("
					+	sql12
					+	" )T1 LEFT OUTER JOIN ("
					+	sql11
					+	" )T2 ON T2.SEIN_CNAME = T1.SEIN_CNAME ";
				
			//��ͨ��Դ����PTHY
			String	sql13="SELECT X.ORG_NAME,COUNT(1) PTHY FROM RSM_OS.RSM_REGISTER_INFO@DQCDBLINK T, MDM_MD.MDM_ORGAN@DQCDBLINK X,RSM_OS.RSM_SCHEDUL_INFO@DQCDBLINK Y"
					+	" WHERE T.ORG_CODE = X.ORG_CODE"
					+	" AND T.SCHEDUL_ID=Y.SCHEDUL_ID"
					+	" AND T.CLINIC_SHIFT=Y.CLINIC_SHIFT"
					+	" AND T.CLINIC_DATE = Y.CLINIC_DATE"
					+	" AND T.ORG_CODE = Y.ORG_CODE"
					+	" AND Y.REG_TYPE_CODE ='1'"
					+	" AND T.CLINIC_DATE >"
					+	"'"+datetime4+"'"
					+	" AND T.CLINIC_DATE <="
					+	"'"+datetime5+"'"
					+	" AND T.REG_USEWAY_CODE IS NOT NULL"
					+	" AND X.ORG_NAME LIKE '%����%' AND   X.ORG_NAME NOT LIKE '%��������%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";
			
			//ר�Һ�Դͳ��ZJHY
			String sql14="SELECT X.ORG_NAME,COUNT(1) ZJHY FROM RSM_OS.RSM_REGISTER_INFO@DQCDBLINK T, MDM_MD.MDM_ORGAN@DQCDBLINK X,RSM_OS.RSM_SCHEDUL_INFO@DQCDBLINK Y"
					+	" WHERE T.ORG_CODE = X.ORG_CODE"
					+	" AND T.SCHEDUL_ID=Y.SCHEDUL_ID"
					+	" AND T.CLINIC_SHIFT=Y.CLINIC_SHIFT"
					+	" AND T.CLINIC_DATE = Y.CLINIC_DATE"
					+	" AND T.ORG_CODE = Y.ORG_CODE"
					+	" AND Y.REG_TYPE_CODE <>'1'"
					+	" AND T.CLINIC_DATE >"
					+	"'"+datetime4+"'"
					+	" AND T.CLINIC_DATE <="
					+	"'"+datetime5+"'"
					+	" AND T.REG_USEWAY_CODE IS NOT NULL"
					+	" AND X.ORG_NAME LIKE '%����%' AND   X.ORG_NAME NOT LIKE '%��������%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";

			//ת���ԴZZHY
			String sql15 = "SELECT X.ORG_NAME, COUNT(1) ZZHY"
					+	" FROM RSM_OS.RSM_REGISTER_INFO@DQCDBLINK T, MDM_MD.MDM_ORGAN@DQCDBLINK X"
					+	" WHERE T.ORG_CODE = X.ORG_CODE"
					+	" AND T.CLINIC_DATE > "
					+	"'"+datetime4+"'"
					+	" AND T.CLINIC_DATE <="
					+	"'"+datetime5+"'"
					+	" AND T.REG_USEWAY_CODE IS NOT NULL"
					+	" AND T.REG_USEWAY_CODE = '02'"
					+	" AND X.ORG_NAME LIKE '%����%' AND   X.ORG_NAME NOT LIKE '%��������%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";

			//ר��ת���Դͳ��ZJZZHY
			String sql16 = "SELECT X.ORG_NAME,COUNT(1) ZJZZHY FROM RSM_OS.RSM_REGISTER_INFO@DQCDBLINK T, MDM_MD.MDM_ORGAN@DQCDBLINK X,RSM_OS.RSM_SCHEDUL_INFO@DQCDBLINK Y"
					+	" WHERE T.ORG_CODE = X.ORG_CODE"
					+	" AND T.SCHEDUL_ID=Y.SCHEDUL_ID"
					+	" AND T.CLINIC_SHIFT=Y.CLINIC_SHIFT"
					+	" AND T.CLINIC_DATE = Y.CLINIC_DATE"
					+	" AND T.ORG_CODE = Y.ORG_CODE"
					+	" AND Y.REG_TYPE_CODE <>'1'"
					+	" AND T.CLINIC_DATE >"
					+	"'"+datetime4+"'"
					+	" AND T.CLINIC_DATE <="
					+	"'"+datetime5+"'"
					+	" AND T.REG_USEWAY_CODE = '02'"
					+	" AND X.ORG_NAME LIKE '%����%' AND   X.ORG_NAME NOT LIKE '%��������%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";

			sql6 = "SELECT T1.*,T2.ZJZZHY"
						+	" FROM (SELECT T1.ORG_NAME,T1.ZZHY,T2.PTHY,T2.ZJHY FROM ("
						+ 	sql15
						+ 	" )T1"
						+	" left join (select T1.ORG_NAME, T1.PTHY, T2.ZJHY FROM ("
						+ 	sql13
						+ 	" )T1"
						+	" LEFT JOIN ("
						+ 	sql14
						+ 	" )T2"
						+	" ON T1.ORG_NAME=T2.ORG_NAME)T2 ON  T1.ORG_NAME = T2.ORG_NAME)T1"
						+	" LEFT JOIN ("
						+ 	sql16
						+ 	" )T2"
						+	" ON T1.ORG_NAME=T2.ORG_NAME";
				System.out.println("���Ե�.......");
			//��ȡ�����
			List<Map<String,Object>> datalist1 = BaseDaoDemo.findListBySql(sql1);
			List<Map<String,Object>> datalist2 = BaseDaoDemo.findListBySql(sql2);
			List<Map<String,Object>> datalist3 = BaseDaoDemo.findListBySql(sql3);
			List<Map<String,Object>> datalist4 = BaseDaoDemo.findListBySql(sql4);
			List<Map<String,Object>> datalist5 = BaseDaoDemo.findListBySql(sql5);
			List<Map<String,Object>> datalist6 = BaseDaoDemo.findListBySql(sql6);
			List<Map<String,Object>> datalist7 = BaseDaoDemo.findListBySql(sql7);
			//����sheet��ͷ
			String[] header1 = { "��������", "��������", "ʧ�ܴ�������)","ƽ����ʱ���룩" ,"ִ���ܴ������Σ�","�ɹ���" };
			String[] header2 = { "��������","��������","ʧ�ܴ�������)","ƽ����ʱ����)","ԭ��" };
			String[] header3 = { "��������","��������","�ɹ���������)","ƽ����ʱ����)" };
			String[] header4 = { "��������","��������","ִ���ܴ���","ƽ����ʱ" };
			String[] header5 = { "��������","��������","����" };
			String[] header6 = { "��������",	"��ͨ��Դ��","ר�Һ�Դ��",	"ת���Դ��	","ר��ת���Դ��","��Դ����","ת���Դռ��"
					,"ר��ת���Դռ��"};
			String[] header7 = { "��������","��������","����" };
			ExportExcelUtils eeu = new ExportExcelUtils();
			HSSFWorkbook workbook = new HSSFWorkbook();

			try {
				eeu.exportExcel(workbook, 0, "�ſ�", header1, datalist1);
				eeu.exportExcel(workbook, 1, "ʧ��", header2, datalist2);
				eeu.exportExcel(workbook, 2, "�ɹ�", header3, datalist3);
				eeu.exportExcel(workbook, 3, "���������������", header4, datalist4);
				eeu.exportExcel(workbook, 4, "�ɼ���������", header5, datalist5);
				eeu.exportExcel(workbook, 5, "�Ű��Դ���"+"("+datetime6+"-"+datetime5+")", header6, datalist6);
				eeu.exportExcel(workbook, 6, "��ѯ������־���", header7, datalist7);
			} catch (Exception e) {
				e.printStackTrace();
			}
	}
}
