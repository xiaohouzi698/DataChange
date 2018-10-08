package ExcelGetUrl;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author 何景辉
 * @date 2018/8/31 16:40
 */
public class Main {

    public void getInfo() throws Exception {

        File file = new File("C:\\Users\\maxin\\Desktop\\supplier-uploads.xls");
        if (!file.exists())
            System.out.println("文件不存在");
        try {
            //1.读取Excel的对象
            // POIFSFileSystem poifsFileSystem = new POIFSFileSystem(new FileInputStream(file));
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(new FileInputStream("C:\\Users\\maxin\\Desktop\\supplier-uploads.xls"));
            //2.Excel工作薄对象
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
            HSSFWorkbook hssfWorkbook2 = new HSSFWorkbook();

            //3.Excel工作表对象
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            //总行数
            int rowLength = hssfSheet.getLastRowNum()+1;
            //4.得到Excel工作表的行
            HSSFRow hssfRow = hssfSheet.getRow(0);
            //总列数
            int colLength = hssfRow.getLastCellNum();
            //得到Excel指定单元格中的内容
            HSSFCell hssfCell = hssfRow.getCell(0);
            //得到单元格样式
            CellStyle cellStyle = hssfCell.getCellStyle();


            //写入操作
            HSSFSheet sheet2 = hssfWorkbook2.createSheet("new");

            // HSSFWorkbook wb = new HSSFWorkbook();


            CellStyle cellStyle2 =hssfWorkbook2.createCellStyle();
            cellStyle2.setFillForegroundColor(HSSFColor.SKY_BLUE.index);




            for (int i = 0; i < rowLength; i++) {
                //获取Excel工作表的行
                HSSFRow hssfRow1 = hssfSheet.getRow(i);

                HSSFRow hssfRow2 = hssfSheet.createRow((int) 0);

                hssfRow2 = sheet2.createRow((int) i + 1);

                for (int j = 0; j < colLength; j++) {
                    //获取指定单元格
                    if (hssfRow1 != null){
                        HSSFCell hssfCell1 = hssfRow1.getCell(j);
                    HSSFCell hssfCell2 = hssfRow2.getCell(j);

                    //Excel数据Cell有不同的类型，当我们试图从一个数字类型的Cell读取出一个字符串时就有可能报异常：
                    //Cannot get a STRING value from a NUMERIC cell
                    //将所有的需要读的Cell表格设置为String格式
                    if (hssfCell1 != null) {
                        hssfCell1.setCellType(CellType.STRING);
                    }


                    // hssfCell2.setCellValue(finish);

                    //获取每一列中的值
                    System.out.print("\t" + i + "行" + j + "列" + "\t");
                    // System.out.print(finish);
                    if (j == 0) {
                        String id = hssfCell1.getStringCellValue();
                        hssfRow2.createCell((short) j).setCellValue(id);

                    } else {
                        String str = hssfCell1.getStringCellValue();
                        String rgex = "/uploads(.*?)g';";
                        String second = (new Main()).getSubUtilSimple(str, rgex);
                        String finish = "http://img1.yikelian.net/uploads" + second + "g";
                        hssfRow2.createCell((short) j).setCellValue(finish);
                        //  hssfCell2.setCellValue(finish);
                    }
                    // System.out.print(hssfCell1.getStringCellValue()+"\n");


                }
                }

                try
                {
                    FileOutputStream fout = new FileOutputStream("C:\\Users\\maxin\\Desktop\\supplier-uploads1.xls");
                    //FileOutputStream fout = new FileOutputStream("E:/students.xls");
                    hssfWorkbook2.write(fout);
                    fout.close();
                }
                catch (Exception e)
                {
                    e.printStackTrace();
                }



                if(i==rowLength-1){
                    System.out.println("\n"+"finish");
                }
                System.out.println("\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    /**
     * 正则表达式匹配两个指定字符串中间的内容
     * @param soap
     * @return
     */
    public List<String> getSubUtil(String soap, String rgex){
        List<String> list = new ArrayList<String>();
        Pattern pattern = Pattern.compile(rgex);// 匹配的模式
        Matcher m = pattern.matcher(soap);
        while (m.find()) {
            int i = 1;
            list.add(m.group(i));
            i++;
        }
        return list;
    }

    /**
     * 返回单个字符串，若匹配到多个的话就返回第一个，方法与getSubUtil一样
     * @param soap
     * @param rgex
     * @return
     */
    public String getSubUtilSimple(String soap,String rgex){
        Pattern pattern = Pattern.compile(rgex);// 匹配的模式
        Matcher m = pattern.matcher(soap);
        while(m.find()){
            return m.group(1);
        }
        return "";
    }



    public static void main (String[]args) throws Exception {
        new Main().getInfo();


        //String a="abcdefg<img src=\"0001.jpg\"/>hijklmnopq";
        //String b=a.replaceAll("<img[^/>]*/>","替换标签");用正则表达式
        //System.out.println(b);

        System.out.println("Hello World!");








    }



}
