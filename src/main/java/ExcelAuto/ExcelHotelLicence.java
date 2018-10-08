package ExcelAuto;

import ExcelGetUrl.Main;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * @author 操作之前清空supplier表中的空白行
 * @date 2018/9/5 11:32
 */
public class ExcelHotelLicence {

    public static void main(String[] args) throws IOException,
            InvalidFormatException {
        String excelPath = "C:\\Users\\maxin\\Desktop\\supplier-uploads.xlsx";

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

                String outPath = "C:\\Users\\maxin\\Desktop\\licence_hotel_supplier1.xlsx";
                Workbook wb2 = null;
                wb2 = new XSSFWorkbook();
                Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
                // 循环写入行数据
                Row row0 = (Row) sheet1.createRow(0);
                Cell cell0 = row0.createCell(0);
                Cell cell1 = row0.createCell(1);

                cell0.setCellValue("1");
                cell1.setCellValue("2");


                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行

                    System.out.println("rIndex: " + rIndex);
                    Row row = sheet.getRow(rIndex);

                    if (row != null) {

                        Cell cellX1 = row.getCell(1); //获取3列数据

                        if(cellX1!=null) {
                            String str = cellX1.getStringCellValue();
                            String rgex = "/uploads(.*?)g';";

                            String second = (new Main()).getSubUtilSimple(str, rgex);
                            String finish = "http://img1.yikelian.net/uploads" + second + "g";


                            System.out.println("行数: " + rIndex);
                            Row rowrIndex = (Row) sheet1.createRow(rIndex);
                            Cell cellx0 = rowrIndex.createCell(2);
                            Cell cellx1 = rowrIndex.createCell(2);
                            Cell cellx2 = rowrIndex.createCell(2);
                            cellx0.setCellValue(finish);

                        }
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
