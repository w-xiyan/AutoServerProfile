package util;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;


public class GetTime {
	   /**
     * ��ȡdata�����
     * @param datatime
     * @return datatime
     * @author XIYAN
     */
	//��ȡǰһ��0��ʱ��
	public static String getStartTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -24);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//�������ڸ�ʽ
		return  df.format(data);
	}
	
	//��ȡ����0��ʱ��
	public static String getnowEndTime() {
		Calendar todayEnd = Calendar.getInstance();
		todayEnd.set(Calendar.HOUR_OF_DAY, 0);
		todayEnd.set(Calendar.MINUTE, 0);
		todayEnd.set(Calendar.SECOND, 0);
		todayEnd.set(Calendar.MILLISECOND,0);
		Date data =  todayEnd.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//�������ڸ�ʽ
		return  df.format(data);
	}
	
	//��ȡǰһ��ʱ���ʽyyyyMMdd
	public static String getYesterdayTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -24);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
		return  df.format(data);
	}
	
	//��ȡ����ʱ���ʽyyyyMMdd
	public static String getTodayTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, 0);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
		return  df.format(data);
	}
	//��ȡ���ո�ʽyyyyMMdd
	public static String getTomorrowTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, 24);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
		return  df.format(data);
	}
	
	//��ȡ��һ��ʱ���ʽyyyyMMdd
	public static String getWeekDayTime() {
		Calendar calendar = Calendar.getInstance();
        Date date = new Date(System.currentTimeMillis());
        calendar.setTime(date);
        calendar.add(Calendar.WEEK_OF_YEAR, 1);
        date = calendar.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
		return  df.format(date);
	}
		
	//��ȡ3��ǰʱ��
	public static String getBeforeThreeTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -72);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//�������ڸ�ʽ
		return  df.format(data);
		}
	//��ȡ3��ǰʱ��yyyyMMdd
	public static String getBeforeThreeToTime() {
		Calendar todayStart = Calendar.getInstance();
		todayStart.set(Calendar.HOUR_OF_DAY, -72);
		todayStart.set(Calendar.MINUTE, 0);
		todayStart.set(Calendar.SECOND, 0);
		todayStart.set(Calendar.MILLISECOND, 0);
		Date data =  todayStart.getTime();
		SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
		return  df.format(data);
			}

	
	
}
