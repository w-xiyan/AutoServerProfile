package util;
 
import java.io.File;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
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
	 * @Description: ����Excel�ķ���
	 * @author: XIYAN@2018-07-03
	 * @param workbook 
	 * @param sheetNum (sheet��λ�ã�0��ʾ��һ�������еĵ�һ��sheet)
	 * @param sheetTitle  ��sheet�����ƣ�
	 * @param headers    ������ı��⣩
	 * @param result   ����������ݣ�
	 * @param out  ���������
	 * @throws Exception
	 */
	public void exportExcel(HSSFWorkbook workbook, int sheetNum,
			String sheetTitle, String[] headers, List<Map<String, Object>> result) throws Exception {
		// ����һ������
		HSSFSheet sheet = workbook.createSheet();
		workbook.setSheetName(sheetNum, sheetTitle);
		//���ñ�ͷ��ʽ
		HSSFCellStyle styleHead = workbook.createCellStyle(); 
		// ���ñ���Ĭ���п���Ϊ20���ֽ�
		sheet.setDefaultColumnWidth((short) 20);
		// ����һ����ʽ
		HSSFCellStyle style = workbook.createCellStyle();
		// ������Щ��ʽ
		style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		
		//���þ��и�ʽ
		style.setAlignment(HorizontalAlignment.CENTER);
		// ����һ������
		HSSFFont font = workbook.createFont();
		font.setColor(HSSFColor.BLACK.index);
		font.setFontHeightInPoints((short) 12);
		// ������Ӧ�õ���ǰ����ʽ
		style.setFont(font);
 
		// ָ������Ԫ��������ʾ����ʱ�Զ�����
		style.setWrapText(true);
		
		//���Ϊ��Դ����ͷ��ʽ�������⴦��
		/*if (sheetTitle == "�Ű��Դ���") {
			// ������0�� Ҳ���Ǳ���
			HSSFRow row1 = sheet.createRow((int) 0);
			row1.setHeightInPoints(50);// �豸����ĸ߶�
			// ������1�� Ҳ���Ǳ�ͷ
			HSSFRow row = sheet.createRow((int) 1);
			row.setHeightInPoints(10);// ���ñ�ͷ�߶�
			HSSFCell cell1 = row1.createCell(0);// ���������һ��
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0,
					headers.length-1)); // �ϲ���0����7��
			cell1.setCellValue(titleName); // ���ñ���ֵ
		}*/
		
		// �������������
		HSSFRow row = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			HSSFCell cell = row.createCell((short) i);
			cell.setCellStyle(style);
			HSSFRichTextString text = new HSSFRichTextString(headers[i]);
			cell.setCellValue(text.toString());
		}
		//������Ԫ�񣬲�����ֵ
		for (int i = 0; i < result.size(); i++) 
		{
			row = sheet.createRow((int) i + 1 );//����������������ӵ�3�п�ʼд���ݣ�
			// Ϊ�������������ص��µ�Ԫ����ʽ1 �Զ����� ���¾���
			HSSFCellStyle autotransfer = workbook.createCellStyle();
			autotransfer.setWrapText(true);// �����Զ�����
			autotransfer.setAlignment(HorizontalAlignment.CENTER);
			// Ϊ�������������ص��µ�Ԫ����ʽ2 �Զ����� ���¾�������Ҳ����
			HSSFCellStyle autotransfer2 = workbook.createCellStyle();
			//���������ʽΪ���壬��СΪ11
			HSSFFont font2 = workbook.createFont();
			font2.setFontName("����");
			font2.setColor(HSSFColor.BLACK.index);
			font2.setFontHeightInPoints((short) 11);
			// ������Ӧ�õ���ǰ����ʽ
			autotransfer2.setFont(font2);
			//���õ�Ԫ���Զ�����
			// ���ñ߿�
			autotransfer2.setBorderBottom(BorderStyle.THIN); //�±߿�    
			autotransfer2.setBorderLeft(BorderStyle.THIN);//��߿�    
			autotransfer2.setBorderTop(BorderStyle.THIN);//�ϱ߿�    
			autotransfer2.setBorderRight(BorderStyle.THIN);//�ұ߿�  
			autotransfer2.setAlignment(HorizontalAlignment.LEFT);//�ӡ�
			autotransfer2.setVerticalAlignment(VerticalAlignment.CENTER);//��ֱ
			//��Ԫ����ɫ����
			//HSSFCellStyle colorCellStyle = workbook.createCellStyle();
			//autotransfer2.setBottomBorderColor(HSSFColor.RED.index);
			//autotransfer2.setFillBackgroundColor(HSSFColor.BLUE.index);
			//autotransfer2.setFillPattern(HSSFCellStyle.FINE_DOTS );
			autotransfer2.setFillBackgroundColor(new HSSFColor.RED().getIndex());
			
			//��Ԫ��ֵ
			HSSFCell datacell = null;
			for (int j = 0; j < headers.length; j++) 
			{
				datacell = row.createCell(j);//��ȡÿһ�еı��
				Map<String, Object> map =result.get(i);
				Set<String> st = map.keySet();
				String value = null;
				//����ڲ�ͨsheet�ĵ�Ԫ��ֵ����
				//�ſ�
				if (sheetTitle == "�ſ�") {
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
						vla1 = Float.valueOf(value);//��string����ת��Ϊfloat����
						value = map.get("ALL_COUNT").toString();
						vla2 = Float.valueOf(value);
						float num= 1-(float)vla1/vla2;  
						DecimalFormat df = new DecimalFormat("0.00%");//��ʽ��С��   
						value = df.format(num);//���ص���String���� 
					}
				}
				//ʧ��	
				if (sheetTitle == "ʧ��") {
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
						if (value == null) {
							value = "";
						}
					}
					if (j == 4) {
						value = map.get("ERR_NAME").toString();
					}
				}
				//�ɹ�
				if (sheetTitle == "�ɹ�") {
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
				}
				//���������������
				if (sheetTitle == "���������������") {
					if (j == 0) {
						 value = map.get("NODE_NAME").toString();
					}    
					if (j == 1) {
						value = map.get("SEIN_CNAME").toString();
					}
					if (j == 2) {
						Object obj = map.get("FAILED_COUNT");
						if (obj == null || "".equals(obj)) {
							value = "";
						}else {
							value = obj.toString();
						}
					}
					if (j == 3) {
						Object obj = map.get("AVG_TIME");
						if (obj == null || "".equals(obj)) {
							value = "";
						}else {
							value = obj.toString();
						}
					}
				}
				//�ɼ�����������
				if (sheetTitle.contains("�ɼ���������") ) {
					if (j == 0) {
						value = map.get("SOURCE_ORG_NAME").toString();
					}
					if (j == 1) {
						value = map.get("CLASS_NAME").toString();
					}
					if (j == 2) {
						value = map.get("COUNTNUM").toString();
					}
				}
				//�Ű��Դ���
				if (sheetTitle.contains("�Ű��Դ���")) {
					float val1 = 0;
					float val2 = 0;
					String val3 = "";
					String val4 = "";
					if (j == 0) {
						value = map.get("ORG_NAME").toString();
					}    
					if (j == 1) {
						//��ȡ��ͨ��Դ��
						value = map.get("PTHY").toString();
					}
					if (j == 2) {
						//��ȡר�Һ�Դ��
						value = map.get("ZJHY").toString();
					}
					if (j == 3) {
						//��ȡת���Դ��
						value = map.get("ZZHY").toString();
					}
					if (j == 4) {
						//��ȡר��ת���Դ��
						value = map.get("ZJZZHY").toString();
					}
					if (j == 5) {
						//��Դ����
						val3 = map.get("PTHY").toString();
						val4 = map.get("ZJHY").toString();
						Integer num = Integer.valueOf(val3)+Integer.valueOf(val4) ;
						value = String.valueOf(num);
						
					}
					if (j == 6) {
						//ר��ת���Դռ��
						value = map.get("ZJZZHY").toString();
						val1 = Float.valueOf(value);//��string����ת��Ϊfloat����
						value = map.get("ZJHY").toString();
						val2 = Float.valueOf(value);
						float num= (float)val1/val2;  
						DecimalFormat df = new DecimalFormat("0.00%");//��ʽ��С��   
						value = df.format(num);//���ص���String���� 
					}
					if (j == 7) {
						//ת���Դռ��
						//1.��ȡת���Դ��
						String val5 = map.get("ZZHY").toString();
						Integer num1 = Integer.valueOf(val5);
						//2.��ȡ��Դ����
						val3 = map.get("PTHY").toString();
						val4 = map.get("ZJHY").toString();
						Integer num2 = Integer.valueOf(val3)+Integer.valueOf(val4) ;
						//3.���м���ռ��
						float num= (float)num1/num2;  
						DecimalFormat df = new DecimalFormat("0.00%");//��ʽ��С��   
						value = df.format(num);//���ص���String���� 
					}
				}
				
				//��Ѳ������־���
				if (sheetTitle == "��ѯ������־���") {
					if (j == 0) {
						 value = map.get("DOMAIN_NAME").toString();
					}    
					if (j == 1) {
						value = map.get("CLASS_NAME").toString();
					}
					if (j == 2) {
						value = map.get("COUNTNUM").toString();
					}
				}
				datacell.setCellValue(value);
				//sheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));//�ϲ���Ԫ��
				datacell.setCellStyle(autotransfer2);//���õ�Ԫ��ĸ�ʽ 
				//datacell.setCellStyle(colorCellStyle);//���õ�Ԫ����ɫ
				
		  	  }
			
		}
		// �����������ļ��浽ָ��λ��
		File file = new File("D:\\�����������ܻ�");
		if (!file.exists()) {
			file.mkdir();//�����ļ���
		}
		FileOutputStream fout = null;
		try {
			SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd"); 
			String nowDate = df.format(new Date()); 
		//	System.out.println(nowDate);
			fout = new FileOutputStream("D:\\�����������ܻ�\\�����������ܻ�"+nowDate+".xls");
			workbook.write(fout);
			String str = "����" + sheetTitle + "�ɹ���";
			System.out.println(str);
		} catch (Exception e) {
			e.printStackTrace();
			String str1 = "����" + sheetTitle + "ʧ�ܣ�";
			System.out.println(str1);
		}finally{
			if (fout == null) {
				fout.close();
			}
			
		}
		 
	}
	
	/** 
     * ������ͷ 
     * @param listColumn 
     */  
    /*public void createHead(List<Column> listColumn,HSSFSheet sheetCo){  
         HSSFRow row = sheetCo.createRow(0);    
         HSSFRow row2 = sheetCo.createRow(1);  
           
         for(short i = 0, n = 0; i < listColumn.size(); i++){//i��headers��������n��Excel������    
            HSSFCell cell1 = row.createCell(n);   
            cell1.setCellStyle(styleHead); //���ñ�ͷ��ʽ   
            HSSFRichTextString text = null;    
            List<Column> columns = listColumn.get(i).getListColumn();  
            if(columns != null && columns.size() > 0){//˫����  
                text = new HSSFRichTextString(listColumn.get(i).getContent());   
                sheetCo.addMergedRegion(new Region(0, n, 0, (short) (n + columns.size() -1)));//������һ�д����  
                  
                short tempI = n;   
                for(int j = 0; j < columns.size(); j++){//���ӱ�����ʽ  
                    HSSFCell cellT = row.createCell(tempI++);    
                     cellT.setCellStyle(styleHead);   
                }  
                for(int j = 0; j < columns.size(); j++){  //�����⸴��  
                    HSSFCell cell2 = row2.createCell(n++);    
                     
                    cell2.setCellStyle(styleHead);    
                    cell2.setCellValue(new HSSFRichTextString(columns.get(j).getContent()));     
               }  
            }else{//������  
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