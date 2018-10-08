package ExcelAuto;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.sql.*;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

/**
 * @author 何景辉
 * @date 2018/9/6 10:55
 */
public class ExcelHotelStateAutoDay {


    private static Connection getConn() {
        String driver = "com.mysql.jdbc.Driver";
        String url = "jdbc:mysql://rm-m5e373u96pv24ke79io.mysql.rds.aliyuncs.com:3306/miyidb";
        String username = "miyidb";
        String password = "Ykl123456";
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

        String LAST = "";
        Workbook wb2 = null;
        String excelPath = "C:\\Users\\maxin\\Desktop\\h_state.xlsx";
            try {
                File excel = new File(excelPath);
                if (excel.isFile() && excel.exists()) {   //判断文件是否存在

                    String[] split = excel.getName().split("\\.");  //.是特殊字符，需要转义！！！！！
                    Workbook wb;
                    //根据文件后缀（xls/xlsx）进行判断
                    if ("xls".equals(split[1])) {
                        FileInputStream fis = new FileInputStream(excel);   //文件流对象
                        wb = new HSSFWorkbook(fis);
                    } else if ("xlsx".equals(split[1])) {
                        wb = new XSSFWorkbook(excel);
                    } else {
                        System.out.println("文件类型错误!");
                        return;
                    }

                    //开始解析
                    Sheet sheet = wb.getSheetAt(0);     //读取sheet 0

                    int firstRowIndex = sheet.getFirstRowNum() + 1;   //第一行是列名，所以不读
                    int lastRowIndex = sheet.getLastRowNum();
                    System.out.println("firstRowIndex: " + firstRowIndex);
                    System.out.println("lastRowIndex: " + lastRowIndex);
                    int number = 1;


                    wb2 = new XSSFWorkbook();
                    Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
                    // 循环写入行数据
                    Row row0 = (Row) sheet1.createRow(0);


                    for (int n = 69; n < 365; n++) {

                        Date dayNow = new Date();   //当前时间
                        Date DayNumber = new Date();

                        Calendar calendar = Calendar.getInstance(); //得到日历
                        calendar.setTime(dayNow);//把当前时间赋给日历
                        calendar.add(Calendar.DAY_OF_MONTH, n);  //设置为前一天
                        DayNumber = calendar.getTime();   //得到前一天的时间

                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd"); //设置时间格式
                        String defaultDayNumber = sdf.format(DayNumber);    //格式化前一天
                        String defaultdayNow = sdf.format(dayNow); //格式化当前时间
                        String day1Now = defaultDayNumber.replaceAll("/", "");

                        System.out.println("第" + n + "天的时间是：" + defaultDayNumber);

                        LAST = day1Now;
                        Connection conn = getConn();
                        int i = 0;
                        String sql = "INSERT INTO h_state VALUES (?,?,?,?,?,?,NULL,?,NULL,NULL,NULL,?,?,?,?,?,?);";
                        PreparedStatement pstmt = null;
                        
                        for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行

                            Row row = sheet.getRow(rIndex);


                            if (row != null) {

                                XSSFCell cellX0 = (XSSFCell) row.getCell(0);
                                cellX0.setCellType(CellType.STRING);
                                XSSFCell cellX1 = (XSSFCell) row.getCell(1);
                                cellX1.setCellType(CellType.STRING);

                                Cell cellX2 = row.getCell(2);
                                XSSFCell cellX3 = (XSSFCell) row.getCell(3);
                                cellX3.setCellType(CellType.STRING);

                                XSSFCell cellX4 = (XSSFCell) row.getCell(4);
                                cellX4.setCellType(CellType.STRING);

                                XSSFCell cellX5 = (XSSFCell) row.getCell(5);
                                cellX5.setCellType(CellType.STRING);

                                XSSFCell cellX11 = (XSSFCell) row.getCell(11);
                                cellX11.setCellType(CellType.STRING);

                                XSSFCell cellX12 = (XSSFCell) row.getCell(12);
                                cellX12.setCellType(CellType.STRING);


                                String hotel_id = cellX0.getStringCellValue();
                                String price_id = cellX1.getStringCellValue();

                                String str = cellX3.getStringCellValue();
                                String breakfast = cellX4.getStringCellValue();
                                String num = cellX5.getStringCellValue();

                                String sale_price = cellX11.getStringCellValue();
                                String base_price = cellX12.getStringCellValue();


                                System.out.println("行数: " + number);


                                String datenow = defaultDayNumber;

                                number++;



                                try {
                                    pstmt = (PreparedStatement) conn.prepareStatement(sql);
                                    pstmt.setString(1, hotel_id);
                                    pstmt.setString(2, price_id);
                                    pstmt.setString(3, datenow);
                                    pstmt.setString(4, breakfast);
                                    pstmt.setString(5, num);
                                    pstmt.setString(6, "1");

                                    pstmt.setString(7, "1");

                                    pstmt.setString(8, sale_price);
                                    pstmt.setString(9, base_price);
                                    pstmt.setString(10, "0");
                                    pstmt.setString(11, "1");
                                    pstmt.setString(12, "1");
                                    pstmt.setString(13, "1");
                                    System.out.println("n="+n+"sql:  " + sql);
                                    i = pstmt.executeUpdate();

                                } catch (SQLException e) {
                                    e.printStackTrace();
                                }


                            }

                        }

                        pstmt.close();
                        conn.close();

                    }

                }

            } catch (Exception e) {
                e.printStackTrace();
            }
        }

}
