package ExcelAuto;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author 何景辉
 * @date 2018/9/21 10:08
 */
public class ExcelTurnDate {
    public static void main(String[] args) throws IOException,
            InvalidFormatException {
        String excelPath = "C:\\Users\\Administrator\\Desktop\\sline_hotel.xlsx";

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

                String outPath = "C:\\Users\\Administrator\\Desktop\\date_TEST_sline_hotel_DATE_1.xlsx";
                Workbook wb2 = null;
                wb2 = new XSSFWorkbook();
                Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
                // 循环写入行数据
                Row row0 = (Row) sheet1.createRow(0);
                Cell cell0 = row0.createCell(0);
                Cell cell1 = row0.createCell(1);
                Cell cell2 = row0.createCell(2);


                cell0.setCellValue("id");
                cell1.setCellValue("phone");
                cell2.setCellValue("mobile");



                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行

                    Row row = sheet.getRow(rIndex);

                    if (row != null) {

                        XSSFCell cellX0 = (XSSFCell) row.getCell(0);
                        cellX0.setCellType(CellType.STRING);

                        XSSFCell cellX1 = (XSSFCell) row.getCell(14);
                        if (cellX1 != null) {

                            cellX1.setCellType(CellType.STRING);

                            String id = cellX0.getStringCellValue();
                            long addtime = Integer.valueOf(cellX1.getStringCellValue()).longValue();;
                            String date;
                            date = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format(new Date(addtime*1000));

                            System.out.println("行数: " + rIndex);
                            Row rowrIndex = (Row) sheet1.createRow(rIndex);
                            Cell cellx0 = rowrIndex.createCell(0);
                            Cell cellx1 = rowrIndex.createCell(1);


                            cellx0.setCellValue(id);
                            cellx1.setCellValue(date);



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
