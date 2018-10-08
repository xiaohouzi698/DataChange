package ExcelAuto;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * @author 何景辉
 * @date 2018/9/13 20:22
 */
public class ExcelGetEmployee {
    public static void main(String[] args) throws IOException,
            InvalidFormatException {

        //获取人员数组

        String excelPath2 = "C:\\Users\\maxin\\Desktop\\employee.xlsx";
        String arr[] = new String[84];
        try {
            File excel2 = new File(excelPath2);
            if (excel2.isFile() && excel2.exists()) {   //判断文件是否存在

                String[] split2 = excel2.getName().split("\\.");  //.是特殊字符，需要转义！！！！！
                Workbook wb2;
                //根据文件后缀（xls/xlsx）进行判断
                if ("xls".equals(split2[1])) {
                    FileInputStream fis = new FileInputStream(excel2);   //文件流对象
                    wb2 = new HSSFWorkbook(fis);
                } else if ("xlsx".equals(split2[1])) {
                    wb2 = new XSSFWorkbook(excel2);
                } else {
                    System.out.println("文件类型错误!");
                    return;
                }

                //开始解析
                Sheet sheet2 = wb2.getSheetAt(1);     //读取sheet 0

                int firstRowIndex2 = sheet2.getFirstRowNum() + 1;   //第一行是列名，所以不读
                int lastRowIndex2 = sheet2.getLastRowNum();
                System.out.println("firstRowIndex: " + firstRowIndex2);
                System.out.println("lastRowIndex: " + lastRowIndex2);

                for (int rIndex = 0; rIndex <=lastRowIndex2; rIndex++) {   //遍历行

                    Row row = sheet2.getRow(rIndex);
                    if (row != null) {

                        XSSFCell cellX1 = (XSSFCell) row.getCell(0);
                        if (cellX1 != null) {
                            cellX1.setCellType(CellType.STRING);

                            String name = cellX1.getStringCellValue();
                            arr[rIndex] = name;


                        }
                    }
                }

            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


        String excelPath = "C:\\Users\\maxin\\Desktop\\employee.xlsx";

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

                String outPath = "C:\\Users\\maxin\\Desktop\\TEST_sline_employee_test_3.xlsx";
                Workbook wb2 = null;
                wb2 = new XSSFWorkbook();
                Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
                // 循环写入行数据
                Row row0 = (Row) sheet1.createRow(0);
                Cell cell0 = row0.createCell(0);
                Cell cell1 = row0.createCell(1);

                cell0.setCellValue("id");
                cell1.setCellValue("name");


                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行

                    Row row = sheet.getRow(rIndex);

                    if (row != null) {

                        XSSFCell cellX0 = (XSSFCell) row.getCell(0);
                        cellX0.setCellType(CellType.STRING);

                        XSSFCell cellX1 = (XSSFCell) row.getCell(1);
                        Row rowrIndex = (Row) sheet1.createRow(rIndex);
                        Cell cellx0 = rowrIndex.createCell(0);
                        Cell cellx1 = rowrIndex.createCell(1);
                        if (cellX0 != null) {
                            if (cellX1 != null) {

                            String id = cellX0.getStringCellValue();
                            String name = cellX1.getStringCellValue();
                                cellX1.setCellType(CellType.STRING);
                                System.out.println("行数: " + rIndex);

                                for (int i = 0; i < 84; i++) {
                                    boolean status = name.contains(arr[i]);
                                    if (status) {
                                        cellx0.setCellValue(id);
                                        String name1 = arr[i];
                                        System.out.println("包 含 " + name1);
                                        cellx1.setCellValue(name1);
                                        // 创建文件流
                                        OutputStream stream = new FileOutputStream(outPath);
                                        // 写入数据
                                        wb2.write(stream);
                                        // 关闭文件流
                                        stream.close();
                                    } else {
                                        //cellx0.setCellValue(id);
                                        //cellx1.setCellValue("");
                                    }

                                    cellx0.setCellValue(id);
                                    cellx1.setCellValue("");

                                }

                            } else {
                                System.out.println("不包含2");

                            }


                        }
                    }
                }




            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
