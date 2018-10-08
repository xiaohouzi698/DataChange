package ExcelAuto;

import ExcelGetUrl.Main;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * @author 何景辉
 * @date 2018/9/5 19:26
 */
public class ExcelHotelState {

    public static void main(String[] args) throws IOException,
            InvalidFormatException {

        for (int n = 0; n < 365; n++) {
            String excelPath = "C:\\Users\\maxin\\Desktop\\h_state.xlsx";
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

            System.out.println("生成的时间是：" + defaultdayNow);


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

                    String outPath = "C:\\Users\\maxin\\Desktop\\STATE\\TEST_h_state_" + day1Now + ".xlsx";
                    Workbook wb2 = null;
                    wb2 = new XSSFWorkbook();
                    Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
                    // 循环写入行数据
                    Row row0 = (Row) sheet1.createRow(0);
                    Cell cell0 = row0.createCell(0);
                    Cell cell1 = row0.createCell(1);
                    Cell cell2 = row0.createCell(2);
                    Cell cell3 = row0.createCell(3);
                    Cell cell4 = row0.createCell(4);
                    Cell cell5 = row0.createCell(5);
                    Cell cell6 = row0.createCell(6);
                    Cell cell7 = row0.createCell(7);
                    Cell cell8 = row0.createCell(8);
                    Cell cell9 = row0.createCell(9);
                    Cell cell10 = row0.createCell(10);
                    Cell cell11 = row0.createCell(11);
                    Cell cell12 = row0.createCell(12);
                    Cell cell13 = row0.createCell(13);
                    Cell cell14 = row0.createCell(14);
                    Cell cell15 = row0.createCell(15);
                    Cell cell16 = row0.createCell(16);
                    Cell cell17 = row0.createCell(17);

                    cell0.setCellValue("hotelid");
                    cell1.setCellValue("price_id");
                    cell2.setCellValue("date");
                    cell3.setCellValue("breakfast");
                    cell4.setCellValue("num");
                    cell5.setCellValue("guarantee");
                    cell6.setCellValue("overtime_set");
                    cell7.setCellValue("unsubscribe");
                    cell8.setCellValue("unsubscribe_policy");
                    cell9.setCellValue("unsubscribe_date_set");
                    cell10.setCellValue("unsubscribe_time_set");
                    cell11.setCellValue("sale_price");
                    cell12.setCellValue("base_price");
                    cell13.setCellValue("commission");
                    cell14.setCellValue("is_open");
                    cell15.setCellValue("can_be_ordered");
                    cell16.setCellValue("bpm_state");


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


                            System.out.println("行数: " + rIndex);
                            Row rowrIndex = (Row) sheet1.createRow(rIndex);
                            Cell cellx0 = rowrIndex.createCell(0);
                            Cell cellx1 = rowrIndex.createCell(1);
                            Cell cellx2 = rowrIndex.createCell(2);
                            Cell cellx3 = rowrIndex.createCell(3);
                            Cell cellx4 = rowrIndex.createCell(4);
                            Cell cellx5 = rowrIndex.createCell(5);
                            Cell cellx6 = rowrIndex.createCell(6);
                            Cell cellx7 = rowrIndex.createCell(7);
                            Cell cellx8 = rowrIndex.createCell(8);
                            Cell cellx9 = rowrIndex.createCell(9);
                            Cell cellx10 = rowrIndex.createCell(10);
                            Cell cellx11 = rowrIndex.createCell(11);
                            Cell cellx12 = rowrIndex.createCell(12);
                            Cell cellx13 = rowrIndex.createCell(13);
                            Cell cellx14 = rowrIndex.createCell(14);
                            Cell cellx15 = rowrIndex.createCell(15);
                            Cell cellx16 = rowrIndex.createCell(16);
                            Cell cellx17 = rowrIndex.createCell(17);



                            String datenow = defaultDayNumber;


                            cellx0.setCellValue(hotel_id);
                            cellx1.setCellValue(price_id);
                            cellx2.setCellValue(datenow);
                            cellx3.setCellValue(breakfast);
                            cellx4.setCellValue(num);
                            cellx5.setCellValue("1");
                            cellx6.setCellValue("");
                            cellx7.setCellValue("1");
                            cellx8.setCellValue("");
                            cellx9.setCellValue("");
                            cellx10.setCellValue("");
                            cellx11.setCellValue(sale_price);
                            cellx12.setCellValue(base_price);
                            cellx13.setCellValue("");
                            cellx14.setCellValue("1");
                            cellx15.setCellValue("1");
                            cellx16.setCellValue("1");


                        }

                    }
                    // 创建文件流
                    OutputStream stream = new FileOutputStream(outPath);
                    // 写入数据
                    wb2.write(stream);
                    // 关闭文件流
                    stream.close();


                } else {
                    System.out.println("找不到指定的文件");
                }
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
    }
}
