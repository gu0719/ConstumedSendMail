package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeUtility;


public class SendEmail {
    Session session;
    private String mailHost        = "";
    private String sender_username = "";
    private String sender_password = "";


    /*
     * 初始化方法
     */
    public SendEmail(String fileName) throws FileNotFoundException {
        InputStream in = new FileInputStream(new File(fileName));
        Properties properties = new Properties();
        try {
            properties.load(in);
            this.mailHost = properties.getProperty("mail.smtp.host");
            this.sender_username = properties.getProperty("mail.sender.username");
            this.sender_password = properties.getProperty("mail.sender.password");
        } catch (IOException e) {
            e.printStackTrace();
        }
        session = Session.getInstance(properties);
        session.setDebug(false);// 开启后有调试信息
    }


    public boolean doSendHtmlEmail(String subject, String sendHtml, String receiveUser, List<File> attachments, Map<String, String> contPic) {
        MimeMessage message = new MimeMessage(session);
        Transport transport = null;

        try {
            // 发件人
            InternetAddress from = new InternetAddress(sender_username);
            message.setFrom(from);

            // 收件人
            if (receiveUser == null && receiveUser.isEmpty()) {
                return false;
            }
            InternetAddress to = new InternetAddress(receiveUser);
            message.setRecipient(Message.RecipientType.TO, to);

            // 邮件主题
            message.setSubject(subject);
            // 内容
            MimeBodyPart mailBody = new MimeBodyPart();
            MimeMultipart mailContent = new MimeMultipart("mixed");
            mailContent.addBodyPart(mailBody);

            MimeMultipart body = new MimeMultipart("related");  //邮件正文也是一个组合体,需要指明组合关系
            mailBody.setContent(body);

            // 邮件正文由html和图片构成
            if (contPic.size() > 0) {
                for (Map.Entry<String, String> entry : contPic.entrySet()) {
                    // 正文图片
                    MimeBodyPart imgPart = new MimeBodyPart();
                    DataSource ds = new FileDataSource(entry.getValue());
                    DataHandler dh = new DataHandler(ds);
                    imgPart.setDataHandler(dh);
                    imgPart.setContentID(entry.getKey());
                    body.addBodyPart(imgPart);
                }
            }

            // html邮件内容
            MimeBodyPart htmlPart = new MimeBodyPart();
            body.addBodyPart(htmlPart);
            MimeMultipart htmlMultipart = new MimeMultipart("alternative");
            htmlPart.setContent(htmlMultipart);
            MimeBodyPart htmlContent = new MimeBodyPart();
            htmlContent.setContent(sendHtml, "text/html;charset=UTF-8");
            htmlMultipart.addBodyPart(htmlContent);

            // 添加附件的内容
            if (attachments != null && attachments.size() > 0) {
                attachments.forEach(attachment -> {
                    try {
                        if ("".equals(attachment)) {
                            return;
                        }
                        MimeBodyPart attachmentBodyPart = new MimeBodyPart();
                        DataSource source = new FileDataSource(attachment);
                        attachmentBodyPart.setDataHandler(new DataHandler(source));

                        // 网上流传的解决文件名乱码的方法，其实用MimeUtility.encodeWord就可以很方便的搞定
                        // 这里很重要，通过下面的Base64编码的转换可以保证你的中文附件标题名在发送时不会变成乱码
                        //sun.misc.BASE64Encoder enc = new sun.misc.BASE64Encoder();
                        //messageBodyPart.setFileName("=?GBK?B?" + enc.encode(attachment.getName().getBytes()) + "?=");

                        //MimeUtility.encodeWord可以避免文件名乱码
                        attachmentBodyPart.setFileName(MimeUtility.encodeWord(attachment.getName()));
                        mailContent.addBodyPart(attachmentBodyPart);
                    } catch (MessagingException | UnsupportedEncodingException e) {
                        e.printStackTrace();
                    }
                });

            }

            // 将multipart对象放到message中
            message.setContent(mailContent);
            // 保存邮件
            message.saveChanges();

            transport = session.getTransport("smtp");
            // smtp验证，就是你用来发邮件的邮箱用户名密码
            transport.connect(mailHost, sender_username, sender_password);
            // 发送
            transport.sendMessage(message, message.getAllRecipients());

            System.out.println("send success!");
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            if (transport != null) {
                try {
                    transport.close();
                } catch (MessagingException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;

    }


}