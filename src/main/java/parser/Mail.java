package parser;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.util.Date;
import java.util.Properties;

public class Mail {

    final String SSL_FACTORY = "javax.net.ssl.SSLSocketFactory";
    final String username = "mailpowerpointpptx@gmail.com"; //changeable
    final String password = "buQGP\\J]m4t5t>b/"; //changeable
    private String mailAdrReceiver;
    private String msgSubject;
    private String mailMsg;

    public Mail(String mailAdrReceiver,String msgSubject,String mailMsg) throws MessagingException {
        this.mailAdrReceiver = mailAdrReceiver;
        this.msgSubject = msgSubject;
        this.mailMsg = mailMsg;
        writeMsg();
    }

    private void writeMsg() throws MessagingException {

        // Get a Properties object
        Properties props = System.getProperties();
        props.setProperty("mail.smtp.host", "smtp.gmail.com");
        props.setProperty("mail.smtp.socketFactory.class", SSL_FACTORY);
        props.setProperty("mail.smtp.socketFactory.fallback", "false");
        props.setProperty("mail.smtp.port", "465");
        props.setProperty("mail.smtp.socketFactory.port", "465");
        props.put("mail.smtp.auth", "true");
        props.put("mail.debug", "true");
        props.put("mail.store.protocol", "pop3");
        props.put("mail.transport.protocol", "smtp");
        props.put("mail.debug", "false");
        Session session = Session.getDefaultInstance(props,
                new Authenticator(){
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }});

        session.setDebug(false);

        // -- Create a new message --
        Message msg = new MimeMessage(session);

        // -- Set the FROM and TO fields --
        msg.setFrom(new InternetAddress("mailpowerpointpptx@gmail.com"));
        msg.setRecipients(Message.RecipientType.TO,
                InternetAddress.parse(mailAdrReceiver,false));
        msg.setSubject(msgSubject);
        msg.setText(mailMsg);
        msg.setSentDate(new Date());
        Transport.send(msg);

    }
}
