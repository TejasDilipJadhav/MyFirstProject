import com.aspose.words.*; //working with word file

import java.io.File; //for working with excel
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//for sending mail
import java.io.IOException;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

class Solution {

    public static String[] readExcel() throws Exception {

        //Add the file path of the excel file that contains the data 
        File file = new File("C:\\Users\\tejas\\OneDrive\\Desktop\\coding\\project\\src\\certificate.xlsx");
        if (file.exists()) {

            FileInputStream fis = new FileInputStream(file);

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            int count = sheet.getRow(0).getPhysicalNumberOfCells();
            

            int rowm = 1;

            String name[] = new String[(count-1) * 3];

            int index = 0;
            while (rowm <count) {
                
                Row row = sheet.getRow(rowm);
                Cell cellEmail = row.getCell(2);
                
                
                
                Cell cellName = row.getCell(0);
                
              
                Cell cellSurName = row.getCell(1);
                
                

                name[index] = cellName.getStringCellValue();
                name[index + count-1] = cellSurName.getStringCellValue();
                name[index + count + count-2] = cellEmail.getStringCellValue();

                rowm++;
                index++;
            }

            wb.close();
            return name;

        }
        String arr[] = { "failed" };
        return arr;
    }

    public static void changeName(String name) throws Exception {

        //Kinldy add the filepath of the word file that contains the design of certificates
        Document doc = new Document("C:\\Users\\tejas\\OneDrive\\Desktop\\coding\\project\\src\\sample.docx");
        // Find and replace text in the document
        doc.getRange().replace("###", name, new FindReplaceOptions(FindReplaceDirection.FORWARD));
        // Save the Word document as pdf
        doc.save("C:\\Users\\tejas\\OneDrive\\Desktop\\coding\\project\\src\\Sending Certificates\\" + name + ".pdf");
    }

    public static void sendEmail(String toEmail, String fileName, String Name) {
        final String username = "from Email Id";
        final String password = "password";
        String fromEmail = "from Email Id";

        Properties properties = new Properties();
        properties.put("mail.smtp.ssl.protocols", "TLSv1.2");
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.ssl.trust", "smtp.gmail.com");
        properties.put("mail.smtp.starttls.enable", "true");
        properties.put("mail.smtp.host", "smtp.gmail.com");
        properties.put("mail.smtp.port", "587");

        Session session = Session.getInstance(properties, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        });
        // Start our mail message
        MimeMessage msg = new MimeMessage(session);
        try {
            msg.setFrom(new InternetAddress(fromEmail));
            msg.addRecipient(Message.RecipientType.TO, new InternetAddress(toEmail));

            //Subject of email
            msg.setSubject("Certificate Of Participation");

            Multipart emailContent = new MimeMultipart();

            // Text body part
            MimeBodyPart textBodyPart = new MimeBodyPart();
            textBodyPart.setText("Congratulations " + Name
                    + "\n\nFor successfully participating in KIA(Know It All) conducted by SEC-GFG. Your participation was of great meaning to us and we really appreciatte your efforts.\n\nHope this certificate helps you build your path towards your dream career.\n\nPFA you certificate of participation\n\nShare your certificates on linkedin and don't forget to tag SEC VIIT and GFG VIIT on your posts\n\nThanks & Regards,\nSEC-VIIT and GFG-VIIT\nBRACT's VIIT Pune.");

            // Attachment body part.
            MimeBodyPart pdfAttachment = new MimeBodyPart();

            //Kidly add location of folder where all certificates will be saved
            pdfAttachment.attachFile(
                    "C:\\Users\\tejas\\OneDrive\\Desktop\\coding\\project\\src\\Sending Certificates\\" + fileName);

            // Attach body parts
            emailContent.addBodyPart(textBodyPart);
            emailContent.addBodyPart(pdfAttachment);

            // Attach multipart to message
            msg.setContent(emailContent);

            Transport.send(msg);

        } catch (MessagingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    public static void main(String[] args) throws Exception {
        String res[] = readExcel(); // create an array that has all the data from the excel sheet
        int count = res.length / 3;

        //Traversing the created array to verify all names and email ids 
        for(int i=0;i<count;i++)
        {
            System.out.println("Name: "+res[i]+" "+res[i+count]+"  Email: "+res[i+count+count]);
        }
        // traverse the array and edit the certificates accordingly
        for (int i = 0; i < count; i++) {
            changeName(res[i] + " " + res[i + count]);
        }

        // Send the certificates to the specified emails
        for (int i = 0; i < count; i++) {
            sendEmail(res[i + count + count], res[i] + " " + res[i + count] + ".pdf", res[i] + " " + res[i + count]);

        }

        // Acknowledgement that the program ran successfully
        System.out.println("Successfull");

    }

}
