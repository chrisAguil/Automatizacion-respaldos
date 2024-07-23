import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.security.Security;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;
import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.search.ReceivedDateTerm;

public class MarcarAltex {
    public static void main(String[] args) {
        Security.setProperty("jdk.tls.disabledAlgorithms", ""); // Consider specifying algorithms to disable if necessary
        Properties properties = new Properties();
        try (FileInputStream fileInputStream = new FileInputStream("configMarcado.conf")) {
            properties.load(fileInputStream);
        } catch (IOException e) {
            e.printStackTrace();
            return; // Exit if configuration cannot be loaded
        }

        String host = properties.getProperty("host");
        String username = properties.getProperty("username");
        String password = properties.getProperty("password");
        String[] asunto = properties.getProperty("asunto").split(",");
        String[] palabraClave = properties.getProperty("palabraClave").split(",");
        String[] marcado = properties.getProperty("marcado").split(",");
        String fechaToday = properties.getProperty("fechaToday");
        String fechaC = properties.getProperty("fecha");

        Properties sessionProperties = new Properties();
        sessionProperties.put("mail.store.protocol", "imaps");

        try {
            Session session = Session.getDefaultInstance(sessionProperties, null);
            Store store = session.getStore("imaps");
            store.connect(host, username, password);

            SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
            Date fecha = formato.parse(fechaC);
            if (fechaToday.equals("Si")) {
                Calendar calendar = Calendar.getInstance();
                calendar.add(Calendar.DATE, -1);
                fecha = calendar.getTime();
            }

            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);

            String fechaString = formato.format(fecha);
            ReceivedDateTerm receivedDateTerm = new ReceivedDateTerm(ReceivedDateTerm.EQ, fecha);
            Message[] messages = inbox.search(receivedDateTerm);

            System.out.println("Comenzando a buscar en buzon con fecha " + fechaString);
            for (Message message : messages) {
                String subject = message.getSubject();
                for (int j = 0; j < asunto.length; j++) {
                    if (subject != null && subject.contains(asunto[j])) {
                        Object content = message.getContent();
                        if (content instanceof String) {
                            String body = (String) content;
                            if (body.contains(palabraClave[j])) {
                                System.out.println("Si fue realizado el respaldo de " + asunto[j] + " con la palabra clave " + palabraClave[j]);
                                marcado[j] = "Si";
                            }
                        } else if (content instanceof Multipart) {
                            Multipart multiPart = (Multipart) content;
                            String body = getTextFromMultipart(multiPart);
                            if (body.contains(palabraClave[j])) {
                                System.out.println("Si fue realizado el respaldo de " + asunto[j] + " con la palabra clave " + palabraClave[j]);
                                marcado[j] = "Si";
                            }
                        }
                    }
                }
            }

            System.out.println("Termino...");
            StringBuilder contenido = new StringBuilder("[Marcado]\nfecha = " + fechaString + "\nmarcado = ");
            for (int i = 0; i < asunto.length; i++) {
                contenido.append(marcado[i]);
                if (i < asunto.length - 1) {
                    contenido.append(",");
                }
            }

            try (FileWriter fileWriter = new FileWriter("Marcado_altex.lobo")) {
                fileWriter.write(contenido.toString());
            }

            inbox.close(false);
            store.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getTextFromMultipart(Multipart multiPart) throws Exception {
        StringBuilder sb = new StringBuilder();
        for (int partCount = 0; partCount < multiPart.getCount(); partCount++) {
            BodyPart part = multiPart.getBodyPart(partCount);
            if (part.isMimeType("text/plain") || part.isMimeType("text/html")) {
                sb.append(part.getContent().toString());
            }
        }
        return sb.toString();
    }
}