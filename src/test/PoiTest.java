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
				//获取上周五日期 格式为	yyyy-MM-dd HH:mm:ss
				 datetime1 = GetTime.getBeforeThreeTime();
				//当天日期 格式为 yyyy-MM-dd HH:mm:ss
				 datetime2 = GetTime.getnowEndTime();
				//获取上周五日期 格式为	yyyyMMdd
				 datetime3 = GetTime.getBeforeThreeToTime();
				//当天日期 格式为	 yyyyMMdd
				 datetime4 = GetTime.getTodayTime();
				//明日日期 格式为	 yyyyMMdd
				 datetime6 = GetTime.getTomorrowTime();
				//往后一周日期 	格式为yyyyMMdd
				 datetime5 = GetTime.getWeekDayTime();
			}else {
				//昨天日期
				 datetime1 = GetTime.getStartTime();
				//当天日期
				 datetime2 = GetTime.getnowEndTime();
				//昨天日期
				 datetime3 = GetTime.getYesterdayTime();
				//当天日期
				 datetime4 = GetTime.getTodayTime();
				//往后一周日期
				 datetime5 = GetTime.getWeekDayTime();
				//明日日期 格式为	 yyyyMMdd
				 datetime6 = GetTime.getTomorrowTime();
			}
			
			//1.巡检HSB服务sql-1-概况
			String sql1 = "";

			//2 .巡检HSB服务sql-1-失败统计（含原因）
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

			//成功
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

			//4.巡检HSB服务sql-整体服务情况
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
			//5.采集报告数量
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
			
			//7.轮询日志情况
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
			
			//1.0.0巡检HSB服务sql-1-概况
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
			
			//1.0.1巡检HSB服务sql-1 - 总次数
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
				
			//普通号源总数PTHY
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
					+	" AND X.ORG_NAME LIKE '%杭州%' AND   X.ORG_NAME NOT LIKE '%社区卫生%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";
			
			//专家号源统计ZJHY
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
					+	" AND X.ORG_NAME LIKE '%杭州%' AND   X.ORG_NAME NOT LIKE '%社区卫生%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";

			//转诊号源ZZHY
			String sql15 = "SELECT X.ORG_NAME, COUNT(1) ZZHY"
					+	" FROM RSM_OS.RSM_REGISTER_INFO@DQCDBLINK T, MDM_MD.MDM_ORGAN@DQCDBLINK X"
					+	" WHERE T.ORG_CODE = X.ORG_CODE"
					+	" AND T.CLINIC_DATE > "
					+	"'"+datetime4+"'"
					+	" AND T.CLINIC_DATE <="
					+	"'"+datetime5+"'"
					+	" AND T.REG_USEWAY_CODE IS NOT NULL"
					+	" AND T.REG_USEWAY_CODE = '02'"
					+	" AND X.ORG_NAME LIKE '%杭州%' AND   X.ORG_NAME NOT LIKE '%社区卫生%'"
					+	" GROUP BY X.ORG_NAME"
					+	" ORDER BY X.ORG_NAME";

			//专家转诊号源统计ZJZZHY
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
					+	" AND X.ORG_NAME LIKE '%杭州%' AND   X.ORG_NAME NOT LIKE '%社区卫生%'"
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
				System.out.println("请稍等.......");
			//获取结果集
			List<Map<String,Object>> datalist1 = BaseDaoDemo.findListBySql(sql1);
			List<Map<String,Object>> datalist2 = BaseDaoDemo.findListBySql(sql2);
			List<Map<String,Object>> datalist3 = BaseDaoDemo.findListBySql(sql3);
			List<Map<String,Object>> datalist4 = BaseDaoDemo.findListBySql(sql4);
			List<Map<String,Object>> datalist5 = BaseDaoDemo.findListBySql(sql5);
			List<Map<String,Object>> datalist6 = BaseDaoDemo.findListBySql(sql6);
			List<Map<String,Object>> datalist7 = BaseDaoDemo.findListBySql(sql7);
			//设置sheet表头
			String[] header1 = { "机构名称", "服务名称", "失败次数（次)","平均耗时（秒）" ,"执行总次数（次）","成功率" };
			String[] header2 = { "机构名称","服务名称","失败次数（次)","平均耗时（秒)","原因" };
			String[] header3 = { "机构名称","服务名称","成功次数（次)","平均耗时（秒)" };
			String[] header4 = { "机构名称","服务名称","执行总次数","平均耗时" };
			String[] header5 = { "机构名称","报告类型","数量" };
			String[] header6 = { "机构名称",	"普通号源数","专家号源数",	"转诊号源数	","专家转诊号源数","号源总数","转诊号源占比"
					,"专家转诊号源占比"};
			String[] header7 = { "机构名称","服务名称","数量" };
			ExportExcelUtils eeu = new ExportExcelUtils();
			HSSFWorkbook workbook = new HSSFWorkbook();

			try {
				eeu.exportExcel(workbook, 0, "概况", header1, datalist1);
				eeu.exportExcel(workbook, 1, "失败", header2, datalist2);
				eeu.exportExcel(workbook, 2, "成功", header3, datalist3);
				eeu.exportExcel(workbook, 3, "服务整体运行情况", header4, datalist4);
				eeu.exportExcel(workbook, 4, "采集报告数量", header5, datalist5);
				eeu.exportExcel(workbook, 5, "排班号源情况"+"("+datetime6+"-"+datetime5+")", header6, datalist6);
				eeu.exportExcel(workbook, 6, "轮询错误日志情况", header7, datalist7);
			} catch (Exception e) {
				e.printStackTrace();
			}
	}
}
