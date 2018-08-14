package util;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;


public class GetTime {
	   /**
     * 获取data结果集
     * @param datatime
     * @return datatime
     * @author XIYAN
     */
	//获取前一天0点时间
	public static String getStartTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -24);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
		return  df.format(data);
	}
	
	//获取当天0点时间
	public static String getnowEndTime() {
		Calendar todayEnd = Calendar.getInstance();
		todayEnd.set(Calendar.HOUR_OF_DAY, 0);
		todayEnd.set(Calendar.MINUTE, 0);
		todayEnd.set(Calendar.SECOND, 0);
		todayEnd.set(Calendar.MILLISECOND,0);
		Date data =  todayEnd.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
		return  df.format(data);
	}
	
	//获取前一天时间格式yyyyMMdd
	public static String getYesterdayTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -24);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//设置日期格式
		return  df.format(data);
	}
	
	//获取当天时间格式yyyyMMdd
	public static String getTodayTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, 0);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//设置日期格式
		return  df.format(data);
	}
	//获取明日格式yyyyMMdd
	public static String getTomorrowTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, 24);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//设置日期格式
		return  df.format(data);
	}
	
	//获取后一周时间格式yyyyMMdd
	public static String getWeekDayTime() {
		Calendar calendar = Calendar.getInstance();
        Date date = new Date(System.currentTimeMillis());
        calendar.setTime(date);
        calendar.add(Calendar.WEEK_OF_YEAR, 1);
        date = calendar.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//设置日期格式
		return  df.format(date);
	}
		
	//获取3天前时间
	public static String getBeforeThreeTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -72);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
		return  df.format(data);
		}
	//获取3天前时间yyyyMMdd
	public static String getBeforeThreeToTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -72);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//设置日期格式
		return  df.format(data);
			}

	
	
}
