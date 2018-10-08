package Baidu;

import net.sf.json.JSONObject;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.text.DecimalFormat;

/**
 * @author 何景辉
 * @date 2018/9/7 16:19
 */
public class getSubwayCity {
    static String AK = "R96OHCfaKVeusE7qkynOoWTOrhYyGlLD"; // 百度地图密钥

    public static void main(String[] args) {

        String coordinate = getsubwayscity();
        System.out.println(coordinate);
        String [] strs = coordinate.split("[,]");
        String lat1 = strs[0];
        String lng2 = strs[1];

        System.out.println("经度为：" + lat1 + "纬度为：" + lng2);

        double lat = 39.88766;
        double lng = 116.467384;
        String Geocoder = getGeocoder(lat,lng);
        System.out.println( "'xxx：" + Geocoder);



    }

    //坐标：116.418038,39.91979
    //116.467384,39.88766


    // 调用百度地图API根据地址，获取坐标


    public static String getsubwayscity() {

            String url = "http://map.baidu.com/?qt=subwayscity&t=123457788";
            String json = loadJSON(url);
            if (json != null && !"".equals(json)) {
                JSONObject obj = JSONObject.fromObject(json);
                //if ("0".equals(obj.getJSONObject("result").getString("error"))) {
                    String cn_name = obj.getJSONObject("result").getJSONObject("subways_city").getString("cn_name");
                    String code = obj.getJSONObject("result").getJSONObject("subways_city").getString("code");

                    DecimalFormat df = new DecimalFormat("#.######");
                    return df.format(cn_name) + "," + df.format(code);
                //}
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
