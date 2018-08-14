package util;
 
import java.io.File;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
 
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.xssf.usermodel.XSSFRow;
 
public class ExportExcelUtils {
	
	/**
	 * @Title: exportExcel
	 * @Description: 导出Excel的方法
	 * @author: XIYAN@2018-07-03
	 * @param workbook 
	 * @param sheetNum (sheet的位置，0表示第一个表格中的第一个sheet)
	 * @param sheetTitle  （sheet的名称）
	 * @param headers    （表格的标题）
	 * @param result   （表格的数据）
	 * @param out  （输出流）
	 * @throws Exception
	 */
	public void exportExcel(HSSFWorkbook workbook, int sheetNum,
			String sheetTitle, ArrayList<String> headers, List<Map<String, Object>> result) throws Exception {
		// 生成一个表格
		HSSFSheet sheet = workbook.createSheet();
		workbook.setSheetName(sheetNum, sheetTitle);
		//设置表头样式
		HSSFCellStyle styleHead = workbook.createCellStyle(); 
		// 设置表格默认列宽度为20个字节
		sheet.setDefaultColumnWidth((short) 20);
		// 生成一个样式
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式
		style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		
		//设置居中格式
		style.setAlignment(HorizontalAlignment.CENTER);
		// 生成一个字体
		HSSFFont font = workbook.createFont();
		font.setColor(HSSFColor.BLACK.index);
		font.setFontHeightInPoints((short) 12);
		// 把字体应用到当前的样式
		style.setFont(font);
 
		// 指定当单元格内容显示不下时自动换行
		style.setWrapText(true);
		
		
		// 产生表格标题行
		HSSFRow row = sheet.createRow(0);
		for (int i = 0; i < headers.size(); i++) {
			HSSFCell cell = row.createCell((short) i);
			cell.setCellStyle(style);
			HSSFRichTextString text = new HSSFRichTextString(headers.get(i));
			cell.setCellValue(text.toString());
		}
		//创建单元格，并设置值
		for (int i = 0; i < result.size(); i++) 
		{
			row = sheet.createRow((int) i + 1 );//创建所需的行数（从第3行开始写数据）
			// 为数据内容设置特点新单元格样式1 自动换行 上下居中
			HSSFCellStyle autotransfer = workbook.createCellStyle();
			autotransfer.setWrapText(true);// 设置自动换行
			autotransfer.setAlignment(HorizontalAlignment.CENTER);
			// 为数据内容设置特点新单元格样式2 自动换行 上下居中左右也居中
			HSSFCellStyle autotransfer2 = workbook.createCellStyle();
			//设置字体格式为宋体，大小为11
			HSSFFont font2 = workbook.createFont();
			font2.setFontName("宋体");
			font2.setColor(HSSFColor.BLACK.index);
			font2.setFontHeightInPoints((short) 11);
			// 把字体应用到当前的样式
			autotransfer2.setFont(font2);
			//设置单元格自动伸缩
			// 设置边框
			autotransfer2.setBorderBottom(BorderStyle.THIN); //下边框    
			autotransfer2.setBorderLeft(BorderStyle.THIN);//左边框    
			autotransfer2.setBorderTop(BorderStyle.THIN);//上边框    
			autotransfer2.setBorderRight(BorderStyle.THIN);//右边框  
			autotransfer2.setAlignment(HorizontalAlignment.LEFT);//居←
			autotransfer2.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
			//单元格颜色设置
			autotransfer2.setFillBackgroundColor(new HSSFColor.RED().getIndex());
			
			//单元格赋值
			HSSFCell datacell = null;
			for (int j = 0; j < headers.size(); j++) 
			{
				datacell = row.createCell(j);//获取每一行的标记
				Map<String, Object> map =result.get(i);
				Set<String> st = map.keySet();
				String value = null;
				//针对于不通sheet的单元格赋值方法
				//概况
				if (sheetTitle == "概况") {
					float vla1 = 0;
					float vla2 = 0;
					if (j == 0) {
						 value = map.get("NODE_NAME").toString();
					}    
					if (j == 1) {
						value = map.get("SEIN_CNAME").toString();
					}
					if (j == 2) {
						value = map.get("FAILED_COUNT").toString();
					}
					if (j == 3) {
						value = map.get("AVG_TIME").toString();
					}
					if (j == 4) {
						value = map.get("ALL_COUNT").toString();
					}
					if (j == 5) {
						value = map.get("FAILED_COUNT").toString();
						vla1 = Float.valueOf(value);//将string类型转换为float类型
						value = map.get("ALL_COUNT").toString();
						vla2 = Float.valueOf(value);
						float num= 1-(float)vla1/vla2;  
						DecimalFormat df = new DecimalFormat("0.00%");//格式化小数   
						value = df.format(num);//返回的是String类型 
					}
				}

				datacell.setCellValue(value);
				//sheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));//合并单元格
				datacell.setCellStyle(autotransfer2);//设置单元格的格式 
				//datacell.setCellStyle(colorCellStyle);//设置单元格颜色
				
		  	  }
			
		}
		// 第六步，将文件存到指定位置
		File file = new File("D:\\各机构服务总汇");
		if (!file.exists()) {
			file.mkdir();//创建文件夹
		}
		FileOutputStream fout = null;
		try {
			SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd"); 
			String nowDate = df.format(new Date()); 
		//	System.out.println(nowDate);
			fout = new FileOutputStream("D:\\各机构服务总汇\\各机构服务总汇"+nowDate+".xls");
			workbook.write(fout);
			String str = "导出" + sheetTitle + "成功！";
			System.out.println(str);
		} catch (Exception e) {
			e.printStackTrace();
			String str1 = "导出" + sheetTitle + "失败！";
			System.out.println(str1);
		}finally{
			if (fout == null) {
				fout.close();
			}
			
		}
		 
	}
	
	/** 
     * 创建表头 
     * @param listColumn 
     */  
    /*public void createHead(List<Column> listColumn,HSSFSheet sheetCo){  
         HSSFRow row = sheetCo.createRow(0);    
         HSSFRow row2 = sheetCo.createRow(1);  
           
         for(short i = 0, n = 0; i < listColumn.size(); i++){//i是headers的索引，n是Excel的索引    
            HSSFCell cell1 = row.createCell(n);   
            cell1.setCellStyle(styleHead); //设置表头样式   
            HSSFRichTextString text = null;    
            List<Column> columns = listColumn.get(i).getListColumn();  
            if(columns != null && columns.size() > 0){//双标题  
                text = new HSSFRichTextString(listColumn.get(i).getContent());   
                sheetCo.addMergedRegion(new Region(0, n, 0, (short) (n + columns.size() -1)));//创建第一行大标题  
                  
                short tempI = n;   
                for(int j = 0; j < columns.size(); j++){//添加标题样式  
                    HSSFCell cellT = row.createCell(tempI++);    
                     cellT.setCellStyle(styleHead);   
                }  
                for(int j = 0; j < columns.size(); j++){  //给标题复制  
                    HSSFCell cell2 = row2.createCell(n++);    
                     
                    cell2.setCellStyle(styleHead);    
                    cell2.setCellValue(new HSSFRichTextString(columns.get(j).getContent()));     
               }  
            }else{//单标题  
                //sheetCo.setColumnWidth(i, columns.get(i).getContent().getBytes().length*2*256);   
                HSSFCell cell2 = row2.createCell(n);    
                cell2.setCellStyle(styleHead);    
                text = new HSSFRichTextString(listColumn.get(i).getContent());    
                sheetCo.addMergedRegion(new Region(0, n, 1, n));    
                n++;  
            }  
            cell1.setCellValue(text);  
         }  */

 
}	