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

class Participant {
    String firstName;
    String lastName;
    String emailId;

    Participant(String firstName, String lastName, String emailId) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.emailId = emailId;

    }

    static void printList(Participant[] participants) {
        for (int i = 0; i < participants.length; i++) {
            System.out.println("Name: " + participants[i].firstName + " " + participants[i].lastName + " Email: "
                    + participants[i].emailId);
        }
    }
}

class Solution {

    public static Participant[] readExcel() throws Exception {

        // Add the file path of the excel file that contains the data
        File file = new File(
                "/Users/tejas/Desktop/Projects/Participation Certificate Generator and Sender/project/src/certificate.xlsx");
        if (file.exists()) {

            FileInputStream fis = new FileInputStream(file);

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            int count = sheet.getPhysicalNumberOfRows();// Gives total number of rows including the heading rows

            int rowm = 1;

            Participant participants[] = new Participant[count - 1];

            int index = 0;
            while (rowm < count) {

                Row row = sheet.getRow(rowm);
                Cell cellEmail = row.getCell(2);

                Cell cellName = row.getCell(0);

                Cell cellSurName = row.getCell(1);

                participants[index] = new Participant(cellName.getStringCellValue(), cellSurName.getStringCellValue(),
                        cellEmail.getStringCellValue());

                rowm++;
                index++;
            }

            wb.close();
            return participants;

        }
        Participant arr[] = new Participant[0];
        return arr;
    }

    public static void changeName(String name) throws Exception {

        // Kinldy add the filepath of the word file that contains the design of
        // certificates

        Document doc = new Document(
                "/Users/tejas/Desktop/Projects/Participation Certificate Generator and Sender/project/src/sample.docx");
        // Find and replace text in the document
        doc.getRange().replace("###", name, new FindReplaceOptions(FindReplaceDirection.FORWARD));
        // Save the Word document as pdf
        doc.save(
                "/Users/tejas/Desktop/Projects/Participation Certificate Generator and Sender/project/Sending Certificates/"
                        + name + ".pdf");
    }

    public static Session login() {
        final String username = "username";
        final String password = "password";

        Properties properties = new Properties();
        properties.put("mail.smtp.ssl.protocols", "TLSv1.2");
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.ssl.trust", "smtp.gmail.com");
        properties.put("mail.smtp.starttls.enable", "true");
        properties.put("mail.smtp.host", "smtp.gmail.com");
        properties.put("mail.smtp.port", "587");

        // Authenticate and return the session object
        Session session = Session.getInstance(properties, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        });

        return session;
    }

    public static void sendEmail(Session session, String toEmail, String fileName, String Name) {
        String fromEmail = "email";
        // Start our mail message
        MimeMessage msg = new MimeMessage(session);
        try {
            msg.setFrom(new InternetAddress(fromEmail));
            msg.addRecipient(Message.RecipientType.TO, new InternetAddress(toEmail));

            // Subject of email
            msg.setSubject("Certificate Of Participation");

            Multipart emailContent = new MimeMultipart();

            // Text body part
            MimeBodyPart textBodyPart = new MimeBodyPart();
            textBodyPart.setText("Congratulations " + Name
                    + "\n\nFor successfully participating in KIA(Know It All) conducted by SEC-GFG. Your participation was of great meaning to us and we really appreciatte your efforts.\n\nHope this certificate helps you build your path towards your dream career.\n\nPFA you certificate of participation\n\nShare your certificates on linkedin and don't forget to tag SEC VIIT and GFG VIIT on your posts\n\nThanks & Regards,\nSEC-VIIT and GFG-VIIT\nBRACT's VIIT Pune.");

            // Attachment body part.
            MimeBodyPart pdfAttachment = new MimeBodyPart();

            // Kidly add location of folder where all certificates will be saved
            pdfAttachment.attachFile(
                    "/Users/tejas/Desktop/Projects/Participation Certificate Generator and Sender/project/Sending Certificates/"
                            + fileName);

            // Attach body parts
            emailContent.addBodyPart(textBodyPart);
            emailContent.addBodyPart(pdfAttachment);

            // Attach multipart to message
            msg.setContent(emailContent);

            Transport.send(msg);
            System.out.println("Done sending " + Name);

        } catch (MessagingException e) {
            e.printStackTrace();
        } catch (IOException e) {

            e.printStackTrace();
        }

    }

    public static void main(String[] args) throws Exception {
        Participant participants[] = readExcel(); // create an array that has all the data from the excel sheet
        // Traversing the created array to verify all names and email ids
        Participant.printList(participants);

        if (participants.length == 0) {
            System.out.println("Failed");
        }

        // traverse the array and edit the certificates accordingly
        for (int i = 0; i < participants.length; i++) {
            changeName(participants[i].firstName + " " + participants[i].lastName);
        }

        Session session = login();

        // Send the certificates to the specified emails
        for (int i = 0; i < participants.length; i++) {
            sendEmail(session, participants[i].emailId,
                    participants[i].firstName + " " + participants[i].lastName + ".pdf",
                    participants[i].firstName + " " + participants[i].lastName);

        }

        // Acknowledgement that the program ran successfully
        System.out.println("Successfull");

    }

}

// TODO: Change username and password at line 105 and 106. Also change the from
// email at line 127
// TODO: Change template path at line 104, excel file path at line 50 and paths
// at line 94,99 and 148
