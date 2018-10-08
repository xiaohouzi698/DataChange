package ExcelGetLatLng;

import net.sf.json.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.io.FileInputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.text.DecimalFormat;

/**
 * @author 何景辉
 * @date 2018/9/3 9:50
 */

public class GetLatLng {

    static String AK = "R96OHCfaKVeusE7qkynOoWTOrhYyGlLD"; // 百度地图密钥

    public static void main(String[] args) throws IOException,
            InvalidFormatException {
        String excelPath = "C:\\Users\\maxin\\Desktop\\testxxxx.xlsx";

        try {
            //String encoding = "GBK";
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

                String outPath = "C:\\Users\\maxin\\Desktop\\testxxxxx1.xlsx";
                Workbook wb2 = null;
                wb2 = new XSSFWorkbook();
                Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
                // 循环写入行数据
                Row row0 = (Row) sheet1.createRow(0);
                Cell cell0 = row0.createCell(0);
                Cell cell1 = row0.createCell(1);
                cell1.setCellValue("lat");
                cell0.setCellValue("lng");

                lastRowIndex=lastRowIndex+1;
                for (int rIndex = 0; rIndex <= lastRowIndex; rIndex++) {   //遍历行
                    System.out.println("rIndex: " + rIndex);
                    Row row = sheet.getRow(rIndex);
                    Row row1 = (Row) sheet1.createRow(rIndex);

                    if (row != null) {
                        int firstCellIndex = row.getFirstCellNum();
                        int lastCellIndex = row.getLastCellNum();
                        //for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) {   //遍历列
                            //Cell cell = row.getCell(10);
                        Cell cell = row.getCell(0);
                            if (cell != null) {
                                String dom = cell.toString();
                                String coordinate = getCoordinate(dom);
                                if(null != coordinate) {
                                    String[] strs = coordinate.split("[,]");

                                    String lat1 = strs[0];
                                    String lng2 = strs[1];


                                    Row rowrIndex = (Row) sheet1.createRow(rIndex);
                                    Cell cellxx = rowrIndex.createCell(1);
                                    Cell cellyy = rowrIndex.createCell(0);
                                    cellxx.setCellValue(lat1);
                                    cellyy.setCellValue(lng2);
                                    System.out.println("lng2++++++"+lat1+","+lng2);

                                    }
                                }
                                System.out.println(cell.toString());
                            //}
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



    public static boolean write( ) throws Exception {
        String outPath = "C:\\Users\\maxin\\Desktop\\test1.xlsx";
        String fileType = outPath.substring(outPath.lastIndexOf(".") + 1, outPath.length());
        System.out.println(fileType);
// 创建工作文档对象
        Workbook wb2 = null;
        if (fileType.equals("xls")) {
            wb2 = new HSSFWorkbook();
        } else if (fileType.equals("xlsx")) {
            wb2 = new XSSFWorkbook();
        } else {
            System.out.println("您的文档格式不正确！");
            return false;
        }
// 创建sheet对象
        Sheet sheet1 = (Sheet) wb2.createSheet("sheet1");
// 循环写入行数据
        for (int i = 0; i < 5; i++) {
            Row row = (Row) sheet1.createRow(i);
// 循环写入列数据
            for (int j = 0; j < 8; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue("测试" + j);
            }
        }
// 创建文件流
        OutputStream stream = new FileOutputStream(outPath);
// 写入数据
        wb2.write(stream);
// 关闭文件流
        stream.close();
        return true;
    }


    public static String getCoordinate(String address) {
        if (address != null && !"".equals(address)) {
            address = address.replaceAll("\\s*", "").replace("#", "栋");
            String url = "http://api.map.baidu.com/geocoder/v2/?address=" + address + "&output=json&ak=" + AK;
            String json = loadJSON(url);
            if (json != null && !"".equals(json)) {
                JSONObject obj = JSONObject.fromObject(json);
                if ("0".equals(obj.getString("status"))) {
                    double lng = obj.getJSONObject("result").getJSONObject("location").getDouble("lng"); // 经度
                    double lat = obj.getJSONObject("result").getJSONObject("location").getDouble("lat"); // 纬度
                    DecimalFormat df = new DecimalFormat("#.######");
                    return df.format(lat) + "," + df.format(lng);
                }
            }
        }
        return null;
    }


    public static String getGeocoder(double lat , double lng) {
        if (lat != 0 && !"".equals(lat) && lng != 0 && !"".equals(lng)) {

            String url = "http://api.map.baidu.com/geocoder?location=" + lat + "," + lng + "&output=json&ak=" + AK;
            String json = loadJSON(url);
            if (json != null && !"".equals(json)) {
                JSONObject obj = JSONObject.fromObject(json);
                if ("OK".equals(obj.getString("status"))) {
                    String formatted_address = obj.getJSONObject("result").getString("formatted_address");
                    String city = obj.getJSONObject("result").getJSONObject("addressComponent").getString("city");
                    String province = obj.getJSONObject("result").getJSONObject("addressComponent").getString("province");
                    String district = obj.getJSONObject("result").getJSONObject("addressComponent").getString("district");
                    String street = obj.getJSONObject("result").getJSONObject("addressComponent").getString("street");
                    String street_number = obj.getJSONObject("result").getJSONObject("addressComponent").getString("street_number");
                    String cityCode = obj.getJSONObject("result").getString("cityCode");

                    return formatted_address + "," + province + "," + city + "," + district + "," + street + "," + street_number + "," + cityCode;
                }
            }
        }
        return null;
    }


    public static String loadJSON(String url) {
        StringBuilder json = new StringBuilder();
        try {
            URL oracle = new URL(url);
            URLConnection yc = oracle.openConnection();
            BufferedReader in = new BufferedReader(new InputStreamReader(yc.getInputStream(), "UTF-8"));
            String inputLine = null;
            while ((inputLine = in.readLine()) != null) {
                json.append(inputLine);
            }
            in.close();
        } catch (MalformedURLException e) {} catch (IOException e) {}
        return json.toString();
    }

}
