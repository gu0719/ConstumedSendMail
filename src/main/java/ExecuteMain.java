import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
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
        AtomicInteger mailCount = new AtomicInteger();
        String confDir = args[0];
        String fileDir = confDir + "/file";
        File cDir = new File(confDir);

        if (cDir == null && cDir.listFiles().length == 0) {
            System.out.println("配置文件路径有误");
            System.exit(1);
        }
        List<File> cList = Arrays.asList(cDir.listFiles());
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

            String theme = map.get(THEME);
            map.remove(EMAIL_ADDRESS);
            map.remove(ATTACH_DIR);
            map.remove(THEME);
            //根据字段替换值
            for (Map.Entry<String, String> entry : map.entrySet()) {
                //正文
                emailContent = emailContent.replace(entry.getKey(), entry.getValue());
            }
            System.out.println("正文（html格式）：" + emailContent);
            System.out.println("附件数：" + fileList.size());
            //发送邮件
            if (sendEmail.doSendHtmlEmail(theme, emailContent, receiverAddress, fileList)) {
                mailCount.getAndIncrement();
            }
            ;
        });
        System.out.println(mailCount.get());
        System.exit(0);

    }

}
