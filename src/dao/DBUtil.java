package dao;

import java.sql.Connection;
import java.sql.DriverManager;

public class DBUtil {
	private final static String driverName = "oracle.jdbc.driver.OracleDriver";    //Çý¶¯¼ÓÔØ
    private final static String url = "jdbc:oracle:thin:@192.26.28.66:1521:orcl";
    private final static String userName = "dqcdb";
    private final static String pwd = "QWE123asd";


    public static Connection getConnection()
    {
        Connection connection = null;
        try{
            Class.forName(driverName);
            connection = DriverManager.getConnection(url,userName,pwd);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return connection;
    }
}

