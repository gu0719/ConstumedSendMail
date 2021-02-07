package util;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;

public class FileReadUtil {


    /**
     * 解析普通文本文件  流式文件 如txt
     *
     * @param path
     * @return
     */
    public static String readTxt(String path) {
        StringBuilder content = new StringBuilder("");
        try {
            File file = new File(path);
            InputStream is = new FileInputStream(file);
            String encode = getJavaEncode(path);
            InputStreamReader isr = new InputStreamReader(is, encode);
            BufferedReader br = new BufferedReader(isr);
            String str = "";
            while (null != (str = br.readLine())) {
                content.append(str);
            }
            br.close();
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("读取文件:" + path + "失败!");
        }
        return content.toString();
    }

    public static String getJavaEncode(String filePath) {
        String encode = EncodingDetect.getJavaEncode(filePath);
        System.out.println(encode);
        return encode;
    }


}
