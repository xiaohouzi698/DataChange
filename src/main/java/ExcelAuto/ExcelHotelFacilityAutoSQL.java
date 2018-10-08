package ExcelAuto;

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
import java.util.*;

/**
 * @author 何景辉
 * @date 2018/9/6 15:48
 */
public class ExcelHotelFacilityAutoSQL {

    private static int range = 127;
    private static int i = 0;
    private static ArrayList <Integer> originalList = new ArrayList <Integer>();
    private static ArrayList <Integer> result = new ArrayList <Integer>();

    static {
        for (int i = 1; i <= range; i++) {
            originalList.add(i);
        }
    }


    public static int[] randomCommon(int min, int max, int n) {
        if (n > (max - min + 1) || max < min) {
            return null;
        }
        int[] result = new int[n];
        int count = 0;
        while (count < n) {
            int num = (int) (Math.random() * (max - min)) + min;
            boolean flag = true;
            for (int j = 0; j < n; j++) {
                if (num == result[j]) {
                    flag = false;
                    break;
                }
            }
            if (flag) {
                result[count] = num;
                count++;
            }
        }
        return result;
    }

    private static Connection getConn() {
        String driver = "com.mysql.jdbc.Driver";
        String url = "jdbc:mysql://rm-m5e373u96pv24ke79io.mysql.rds.aliyuncs.com:3306/miyidb";
        String username = "miyiuser";
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

/*
        String LAST = "";
        Workbook wb2 = null;
        String excelPath = "C:\\Users\\maxin\\Desktop\\hotel_test_use.xlsx";

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
                int i = 0;
                Connection conn = getConn();


                String sql = "INSERT INTO h_hotel_facility VALUES (?,?,?,NULL,NULL);";
                PreparedStatement pstmt = (PreparedStatement)conn.prepareStatement(sql);
                pstmt = (PreparedStatement) conn.prepareStatement(sql);




                for (int rIndex = 2; rIndex <= lastRowIndex; rIndex++) {   //遍历行

                    Row row = sheet.getRow(rIndex);

                    if (row != null) {

                        Long timeid = System.currentTimeMillis();
                        String time = String.valueOf(timeid);

                        XSSFCell cellX0 = (XSSFCell) row.getCell(0);
                        cellX0.setCellType(CellType.STRING);
                        String hotel_id = cellX0.getStringCellValue();


                        XSSFCell cellX1 = (XSSFCell) row.getCell(7);
                        cellX1.setCellType(CellType.STRING);
                        String starString = cellX1.getStringCellValue();
                        int star = Integer.parseInt(starString);

                        String score=null;


                        try {
                            if (star == 3) {
                                try {

                                    for (int num=0;num<58;num++) {

                                        pstmt.setString(1, hotel_id);

                                        if(num ==0)
                                        {
                                            score="4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "6";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==2) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }else if (num ==1) {
                                            score = "4";
                                            pstmt.setString(2, score);
                                            i = pstmt.executeUpdate();
                                        }

                                    }



                                } catch (SQLException e) {
                                    e.printStackTrace();

                            }
                            } else if (star == 4) {
                                try {
                                    int[] reult4 = randomCommon(1, 127, 82);
                                    for (int rand : reult4) {

                                        String xxx = time + rand;

                                        pstmt.setString(1, hotel_id);
                                        String score = String.valueOf(rand);
                                        pstmt.setString(2, score);
                                        i = pstmt.executeUpdate();
                                        System.out.println("2success: " + rIndex);
                                    }

                                } catch (SQLException e) {
                                    e.printStackTrace();
                                }
                            } else if (star == 5) {
                                try {
                                    int[] reult5 = randomCommon(1, 127, 118);
                                    for (int rand : reult5) {

                                        String xxx = time + rand;

                                        pstmt.setString(1, hotel_id);
                                        String score = String.valueOf(rand);
                                        pstmt.setString(2, score);
                                        i = pstmt.executeUpdate();
                                        System.out.println("2success: " + rIndex);
                                    }

                                } catch (SQLException e) {
                                    e.printStackTrace();
                                }
                            } else {

                                try {
                                    int[] reult2 = randomCommon(1, 127, 18);
                                    for (int rand : reult2) {

                                        String xxx = time + rand;

                                        pstmt.setString(1, hotel_id);
                                        String score = String.valueOf(rand);
                                        pstmt.setString(2, score);
                                        i = pstmt.executeUpdate();
                                        System.out.println("hangshu: " + rIndex);

                                    }

                                } catch (SQLException e) {
                                    e.printStackTrace();
                                }

                            }

                        } catch (Exception e) {
                            e.printStackTrace();
                        }

                    }




                }
                pstmt.close();
                conn.close();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}


*/
    }
}