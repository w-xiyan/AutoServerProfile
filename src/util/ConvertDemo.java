package util;

import java.util.Date; 
import java.text.ParseException; 
import java.text.SimpleDateFormat; 

public class ConvertDemo {
	/** 
	* 日期转换成字符串 
	* @param date 
	* @return str 
	*/ 
	public static String DateToStr(Date date) { 
	  
	   SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); 
	   String str = format.format(date); 
	   return str; 
	} 

	/** 
	* 字符串转换成日期 
	* @param str 
	* @return date 
	*/ 
	public static Date StrToDate(String str) { 
	  
	   SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); 
	   Date date = null; 
	   try { 
	    date = format.parse(str); 
	   } catch (ParseException e) { 
	    e.printStackTrace(); 
	   } 
	   return date; 
	} 
}
