package SQL;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * @author 何景辉
 * @date 2018/9/17 18:12
 */
public class SQLSetComment {

    //INSERT INTO `hjh_dubbo`.`h_comment` (`id`, `star`, `person_id`, `hotel_id`, `layout_id`, `order_Id`, `content`, `review`, `images`, `reply_content`, `state`, `health_score`, `environment_score`, `service_score`, `facilities_score`, `tags`, `create_time`, `praise_num`) VALUES ('153605929650469', '0', '153595339828101', '100023', '153441070121927', '153605921662878', '不满意', NULL, 'http://pdr1j9rox.bkt.clouddn.com/DC2D5887-EE17-4DCC-8D66-37BAF45E5CB51536059295.9980431299284690,http://pdr1j9rox.bkt.clouddn.com/DC2D5887-EE17-4DCC-8D66-37BAF45E5CB51536059295.9679483243894404', NULL, '1', '0', '0', '0', '0', NULL, '2018-09-04 19:08:17', '1');

    private static Connection getConn() {
        String driver = "com.mysql.jdbc.Driver";
        String url = "jdbc:mysql://192.168.1.104:3306/hjh_dubbo";
        String username = "root";
        String password = "ykl123456";
        Connection conn = null;
        try {
            Class.forName(driver); //classLoader,加载对应驱动
            conn = (Connection) DriverManager.getConnection(url, username, password);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return conn;
    }

    public static void main(String[] args) throws IOException,
            InvalidFormatException {

                for (int n = 111169; n < 111170; n++) {

                    Connection conn = getConn();
                    int i = 0;
                    PreparedStatement pstmt = null;
                    String sql = "INSERT INTO h_comment ('id', 'star', 'person_id', 'hotel_id', 'layout_id', 'order_Id', 'content', 'review', 'images', 'reply_content', 'state', 'health_score', 'environment_score', 'service_score', 'facilities_score', 'tags', 'create_time', 'praise_num') VALUES (?, '0', '153595339828101', '100023', '153441070121927', '153605921662878', '不满意', NULL, 'http://pdr1j9rox.bkt.clouddn.com/DC2D5887-EE17-4DCC-8D66-37BAF45E5CB51536059295.9980431299284690,http://pdr1j9rox.bkt.clouddn.com/DC2D5887-EE17-4DCC-8D66-37BAF45E5CB51536059295.9679483243894404', NULL,'1','0','0','0','0',NULL,'2018-09-04 19:08:17','1');";

                    try {

                        for (int m = 111169; m < 141169; m++) {
                            String id = String.valueOf(m);
                            pstmt = (PreparedStatement) conn.prepareStatement(sql);
                            pstmt.setString(1, id);
                            System.out.println("第"+m+"次");
                            i = pstmt.executeUpdate();

                        }
                        pstmt.close();
                        conn.close();
                            } catch (SQLException e) {
                                e.printStackTrace();
                            }

                }



            }





}
