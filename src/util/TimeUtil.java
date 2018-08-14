package util;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

public class TimeUtil {
	/**
     * ��ȡ��ȥ����δ�� �������ڵ���������
     * @param intervals      intervals����
     * @return              ��������
     */
    public static ArrayList<String> getTime(int intervals ) {
        ArrayList<String> pastDaysList = new ArrayList<>();
        ArrayList<String> fetureDaysList = new ArrayList<>();
        for (int i = 0; i <intervals; i++) {
            pastDaysList.add(getPastDate(i));
            fetureDaysList.add(getFetureDate(i+1));
        }
        return fetureDaysList;
    }
 
    /**
     * ��ȡ��ȥ�ڼ��������
     *
     * @param past
     * @return
     */
    public static String getPastDate(int past) {
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DAY_OF_YEAR, calendar.get(Calendar.DAY_OF_YEAR) - past);
        Date today = calendar.getTime();
        SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
        String result = format.format(today);
        //Log.e(null, result);
        return result;
    }
 
    /**
     * ��ȡδ�� �� past �������
     * @param past
     * @return
     */
    public static String getFetureDate(int past) {
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.DAY_OF_YEAR, calendar.get(Calendar.DAY_OF_YEAR) + past);
        Date today = calendar.getTime();
        SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");//�������ڸ�ʽ
        String result = format.format(today);
        //Log.e(null, result);
        return result;

}
}