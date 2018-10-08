package ExcelAuto;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


/**
 * @author 何景辉
 * @date 2018/9/5 9:51
 */
public class ExcelHotelPolicy {



    public static void main(String[] args) throws IOException,
            InvalidFormatException {
        String excelPath = "C:\\Users\\maxin\\Desktop\\hotel_test.xlsx";

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

                String outPath = "C:\\Users\\maxin\\Desktop\\test_hotel_policy.xlsx";
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
                cell0.setCellValue("id");
                cell1.setCellValue("hotelid");
                cell2.setCellValue("policy_title");
                cell3.setCellValue("policy_content");
                cell4.setCellValue("status");

                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行
                    int yIndex=rIndex;

                    System.out.println("rIndex: " + rIndex);
                    Row row = sheet.getRow(rIndex);

                    if (row != null) {
                        int firstCellIndex = row.getFirstCellNum();
                        int lastCellIndex = row.getLastCellNum();

                        Cell cell10 = row.getCell(0);


                            String c10 = cell10.toString();

                            System.out.println("行数: " + yIndex);
                            double hotel_id = Double.parseDouble(c10);

                                for (int num = 1; num < 7; num++) {
                                    if (num == 1) {
                                        yIndex = (yIndex-1)*6 +num;
                                        Row rowrIndex = (Row) sheet1.createRow(yIndex);
                                        Cell cellx0 = rowrIndex.createCell(0);
                                        Cell cellx1 = rowrIndex.createCell(1);
                                        Cell cellx2 = rowrIndex.createCell(2);
                                        Cell cellx3 = rowrIndex.createCell(3);
                                        Cell cellx4 = rowrIndex.createCell(4);
                                        cellx0.setCellValue("400000" + num + rIndex);
                                        cellx1.setCellValue(hotel_id);
                                        cellx2.setCellValue("住离店政策");
                                        cellx3.setCellValue("本酒店支持的最早入住时间为14:00；离店手续需要在14:00前在前台办理");
                                        cellx4.setCellValue("4");
                                        yIndex++;
                                        System.out.println("yIndex "+yIndex+"\n");
                                    } else if (num == 2) {
                                        yIndex = (rIndex-1)*6 +num;
                                        Row rowrIndex = (Row) sheet1.createRow(yIndex);
                                        Cell cellx0 = rowrIndex.createCell(0);
                                        Cell cellx1 = rowrIndex.createCell(1);
                                        Cell cellx2 = rowrIndex.createCell(2);
                                        Cell cellx3 = rowrIndex.createCell(3);
                                        Cell cellx4 = rowrIndex.createCell(4);
                                        cellx0.setCellValue("400000" + num + rIndex);
                                        cellx1.setCellValue(hotel_id);
                                        cellx2.setCellValue("儿童政策");
                                        cellx3.setCellValue("不接受16岁以下未成年人在无监护人陪同的情况下入住；允许客人携带儿童入住，每间客房最多可容纳1名儿童，和成人共用现有床位");
                                        cellx4.setCellValue("4");
                                        yIndex++;
                                        System.out.println("yIndex "+yIndex+"\n");
                                    } else if (num == 3) {
                                        yIndex = (rIndex-1)*6 +num;
                                        Row rowrIndex = (Row) sheet1.createRow(yIndex);
                                        Cell cellx0 = rowrIndex.createCell(0);
                                        Cell cellx1 = rowrIndex.createCell(1);
                                        Cell cellx2 = rowrIndex.createCell(2);
                                        Cell cellx3 = rowrIndex.createCell(3);
                                        Cell cellx4 = rowrIndex.createCell(4);
                                        cellx0.setCellValue("400000" + num + rIndex);
                                        cellx1.setCellValue(hotel_id);
                                        cellx2.setCellValue("酒店服务");
                                        cellx3.setCellValue("本酒店提供自助早餐。早餐费用请前往前台进行咨询。\n" +
                                                "本酒店部分房型提供收费的加床服务。加床费用请前往前台进行咨询。");
                                        cellx4.setCellValue("4");
                                        yIndex++;
                                        System.out.println("yIndex "+yIndex+"\n");
                                    } else if (num == 4) {
                                        yIndex = (rIndex-1)*6 +num;
                                        Row rowrIndex = (Row) sheet1.createRow(yIndex);
                                        Cell cellx0 = rowrIndex.createCell(0);
                                        Cell cellx1 = rowrIndex.createCell(1);
                                        Cell cellx2 = rowrIndex.createCell(2);
                                        Cell cellx3 = rowrIndex.createCell(3);
                                        Cell cellx4 = rowrIndex.createCell(4);
                                        cellx0.setCellValue("400000" + num + rIndex);
                                        cellx1.setCellValue(hotel_id);
                                        cellx2.setCellValue("宠物");
                                        cellx3.setCellValue("本酒店不允许携带宠物入住，敬请谅解。");
                                        cellx4.setCellValue("4");
                                        yIndex++;
                                        System.out.println("yIndex "+yIndex+"\n");
                                    } else if (num == 5) {
                                        yIndex = (rIndex-1)*6 +num;
                                        Row rowrIndex = (Row) sheet1.createRow(yIndex);
                                        Cell cellx0 = rowrIndex.createCell(0);
                                        Cell cellx1 = rowrIndex.createCell(1);
                                        Cell cellx2 = rowrIndex.createCell(2);
                                        Cell cellx3 = rowrIndex.createCell(3);
                                        Cell cellx4 = rowrIndex.createCell(4);
                                        cellx0.setCellValue("400000" + num + rIndex);
                                        cellx1.setCellValue(hotel_id);
                                        cellx2.setCellValue("可支付方式");
                                        cellx3.setCellValue("本酒店支持支付宝支付、微信支付及带有银联标志的银行卡支付");
                                        cellx4.setCellValue("4");
                                        yIndex++;
                                        System.out.println("yIndex "+yIndex+"\n");
                                    } else if(num == 6) {
                                        yIndex = (rIndex-1)*6 +num;
                                        Row rowrIndex = (Row) sheet1.createRow(yIndex);
                                        Cell cellx0 = rowrIndex.createCell(0);
                                        Cell cellx1 = rowrIndex.createCell(1);
                                        Cell cellx2 = rowrIndex.createCell(2);
                                        Cell cellx3 = rowrIndex.createCell(3);
                                        Cell cellx4 = rowrIndex.createCell(4);
                                        cellx0.setCellValue("400000" + num + rIndex);
                                        cellx1.setCellValue(hotel_id);
                                        cellx2.setCellValue("会员政策");
                                        cellx3.setCellValue("平台注册会员享受与酒店会员一致的政策。具体政策内容，请前往前台进行咨询。");
                                        cellx4.setCellValue("4");
                                        yIndex++;
                                        System.out.println("yIndex "+yIndex+"\n");
                                    }


                                }

                            }


                        // System.out.println(cellxx.toString());
                        //}

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
