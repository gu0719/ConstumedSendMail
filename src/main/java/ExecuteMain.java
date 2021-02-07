import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import javax.mail.MessagingException;

import util.ExcelHandleUtil;
import util.FileReadUtil;
import util.SendEmail;

/**
 * 发送邮件自定义
 * 读取文件
 * 包装邮件
 * 发送
 */
public class ExecuteMain {
    private static String EMAIL_ADDRESS  = "$EMAIL_ADDRESS";
    private static String ATTACH_DIR     = "$ATTACH_DIR";
    private static String THEME          = "$THEME";
    private static String propertiesFile = "";
    private static String excelFile      = "";
    private static String modelFile      = "";

    public static void main(String[] args) throws MessagingException, IOException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        AtomicInteger mailCount = new AtomicInteger();
        String confDir = args[0];
        String fileDir = confDir + "/file";
        File cDir = new File(confDir);
        if (cDir == null && cDir.listFiles().length == 0) {
            System.out.println("配置文件路径有误");
            System.exit(1);
        }
        List<File> cList = Arrays.asList(cDir.listFiles());
        List<String> failList = new ArrayList<>();
        List<String> successList = new ArrayList<>();
        cList.forEach(file -> {
            if ("MailServer.properties".equals(file.getName())) {
                propertiesFile = file.getPath();
            }
            if ("Data.xlsx".equals(file.getName())) {
                excelFile = file.getPath();
            }
            if ("mailContent.txt".equals(file.getName())) {
                modelFile = file.getPath();
            }
        });

        SendEmail sendEmail = new SendEmail(propertiesFile);
        //读取excel
        List<Map<String, String>> list = ExcelHandleUtil.getExcelInfo(excelFile);
        Iterator iterator = list.iterator();
        //去除不必要的数据
        while (iterator.hasNext()) {
            Map<String, String> map = (Map<String, String>) iterator.next();
            if (map.get(EMAIL_ADDRESS) == null || "".equals(map.get(EMAIL_ADDRESS))) {
                iterator.remove();
            }
        }
        System.out.println("实际应发送的邮件数：" + list.size());
        System.out.println(list);
        if (list.size() <= 0) {
            System.out.println("读取excel有误，请检查");
            System.exit(1);
        }
        //获取模板
        String mailContentModel = FileReadUtil.readTxt(modelFile);
        //遍历发送邮件
        list.forEach(map -> {
            List<File> fileList = new ArrayList<>();
            String emailContent = mailContentModel;
            String receiverAddress = map.get(EMAIL_ADDRESS);
            String fileName = map.get(ATTACH_DIR);
            if (!"".equals(fileName)) {
                //装配附件
                File dir = new File(fileDir + "/" + fileName);
                if (dir.exists()) {
                    if (dir.isDirectory()) {
                        File[] files = dir.listFiles();
                        if (files != null && files.length > 0) {
                            fileList.addAll(Arrays.asList(files));
                        }
                    } else {
                        fileList.add(dir);
                    }
                }
            }

            String theme = map.get(THEME);
            map.remove(EMAIL_ADDRESS);
            map.remove(ATTACH_DIR);
            map.remove(THEME);
            Map<String, String> contPic = new HashMap<>();
            //根据字段替换值
            for (Map.Entry<String, String> entry : map.entrySet()) {
                if ("".equals(entry.getValue())) {
                    continue;
                }
                if (entry.getKey().contains("$CONT_PIC")) {
                    emailContent = emailContent.replace(entry.getKey(), entry.getKey().replace("$", "") + ".jpg");
                    contPic.put(entry.getKey().replace("$", "") + ".jpg", fileDir + "/" + entry.getValue());
                } else {
                    emailContent = emailContent.replace(entry.getKey(), entry.getValue());

                }

            }
            //发送邮件
            if (sendEmail.doSendHtmlEmail(theme, emailContent, receiverAddress, fileList, contPic)) {
                System.out.println("正文（html格式）：" + emailContent);
                System.out.println("附件数：" + fileList.size());
                mailCount.getAndIncrement();
                successList.add(receiverAddress + "/" + theme);
            } else {
                failList.add(receiverAddress);
            }
        });
        System.out.println("发送成功邮件：" + mailCount.get());
        if (successList.size() > 0) {
            BufferedWriter bw1 = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("./successResult.txt"), StandardCharsets.UTF_8));
            bw1.write(sdf.format(new Date()));
            successList.forEach(successStr -> {
                try {
                    bw1.write(successStr);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
            bw1.write("-----end-----");
            bw1.close();
        }
        if (failList.size() > 0) {
            BufferedWriter bw2 = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("./failResult.txt"), StandardCharsets.UTF_8));
            bw2.write(sdf.format(new Date()));
            failList.forEach(failStr -> {
                try {
                    if (failStr != null) {
                        bw2.write(failStr + "\r\n");
                    } else {
                        bw2.write("发送成功");
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
            bw2.write("-----end-----");
            bw2.close();
        }
        System.exit(0);

    }

}
