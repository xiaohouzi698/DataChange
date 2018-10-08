package GetLocation;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.text.DecimalFormat;
import net.sf.json.JSONObject;

/**
 * @author 何景辉
 * @date 2018/8/31 16:43
 */
public class EntCoordSyncJob {
    static String AK = "R96OHCfaKVeusE7qkynOoWTOrhYyGlLD"; // 百度地图密钥

    public static void main(String[] args) {
        String dom = "三亚市七仙岭旅游度假区";
        String coordinate = getCoordinate(dom);
        System.out.println("'" + dom + "'的经纬度为：" + coordinate);
        String [] strs = coordinate.split("[,]");
        String lat1 = strs[0];
        String lng2 = strs[1];
        //lat1.substring(0, coordinate.indexOf(","));
       // String lng2 = coordinate.substring(coordinate.indexOf(",") + 1);
        System.out.println("经度为：" + lat1 + "纬度为：" + lng2);
        // System.err.println("######同步坐标已达到日配额6000限制，请明天再试！#####");
        double lat = 39.88766;
        double lng = 116.467384;
        String Geocoder = getGeocoder(lat,lng);
        System.out.println( "'xxx：" + Geocoder);



    }

    //坐标：116.418038,39.91979
    //116.467384,39.88766


    // 调用百度地图API根据地址，获取坐标


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

    // 来自stackoverflow的MD5计算方法，调用了MessageDigest库函数，并把byte数组结果转换成16进制
    /*
     * public String MD5(String md5) { try { java.security.MessageDigest md = java.security.MessageDigest .getInstance("MD5"); byte[] array = md.digest(md5.getBytes()); StringBuffer sb = new StringBuffer(); for (int i = 0; i < array.length; ++i) { sb.append(Integer.toHexString((array[i] & 0xFF) | 0x100) .substring(1, 3)); } return sb.toString(); } catch (java.security.NoSuchAlgorithmException e) {
     * } return null; }
     */



}


//http://api.map.baidu.com/geocoder?location=39.88766,116.467384&output=json&key=R96OHCfaKVeusE7qkynOoWTOrhYyGlLD

/*

{
    "status":"OK",
    "result":{
        "location":{
            "lng":116.467384,
            "lat":39.88766
        },
        "formatted_address":"北京市朝阳区劲松南路826号",
        "business":"潘家园,劲松,南磨房",
        "addressComponent":{
            "city":"北京市",
            "direction":"附近",
            "distance":"10",
            "district":"朝阳区",
            "province":"北京市",
            "street":"劲松南路",
            "street_number":"826号"
        },
        "cityCode":131
    }
}

 */