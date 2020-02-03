import com.google.zxing.BarcodeFormat;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Main {

    private final static int QR_WIDTH = 500;
    private final static int QR_HEIGHT = 500;
    private final static String FOLDERNAME = "QR Codes";

    private static void generateQRCodeImage(String text, String path, String name, String matricNo, int groupNo)
            throws WriterException, IOException {
        QRCodeWriter qrCodeWriter = new QRCodeWriter();
        BitMatrix bitMatrix = qrCodeWriter.encode(text, BarcodeFormat.QR_CODE, QR_WIDTH, QR_HEIGHT);

        BufferedImage qrImage = MatrixToImageWriter.toBufferedImage(bitMatrix);
        Graphics2D g = (Graphics2D) qrImage.getGraphics();
        g.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 30));

        //WRAP STRING ACCORDING TO LENGTH
        ArrayList<String> tempStringList = new ArrayList<>();
        FontMetrics fm = g.getFontMetrics();
        wrapString(name, tempStringList, fm);
        tempStringList.add(matricNo);
        tempStringList.add("Group " + groupNo);

//        Calculate space needed for the text
        int totalHeight = fm.getHeight() * tempStringList.size() + 10;
        BufferedImage label = new BufferedImage(QR_WIDTH, totalHeight, BufferedImage.TYPE_INT_ARGB);
        g = (Graphics2D) label.getGraphics();
        g.setRenderingHint(
                RenderingHints.KEY_TEXT_ANTIALIASING,
                RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 30));
        g.setColor(Color.BLACK);
        fm = g.getFontMetrics();
        int x, y;
        int nextLine = 0;
        for (String str : tempStringList) {
            x = (label.getWidth() - fm.stringWidth(str)) / 2;
            y = fm.getHeight() + nextLine;
            g.drawString(str, x, y);
            nextLine += fm.getHeight();
        }

        //Merge QR Image and Label together
        BufferedImage finalImage = new BufferedImage(QR_WIDTH,
                QR_HEIGHT + label.getHeight(), BufferedImage.TYPE_INT_ARGB);
        g = (Graphics2D) finalImage.getGraphics();
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, finalImage.getWidth(), finalImage.getHeight());
        g.drawImage(qrImage, 0, 0, qrImage.getWidth(), qrImage.getHeight(), null);
        g.setColor(Color.BLACK);
        g.drawImage(label, 0, qrImage.getHeight() - 25, label.getWidth(), label.getHeight(), null);
        g.dispose();

        ImageIO.write(finalImage, "png", new File(path));
    }

    private static void wrapString(String tempName, ArrayList<String> tempStringList, FontMetrics fm) {
        if (fm.stringWidth(tempName) > QR_WIDTH - 32) {
            String[] words = tempName.split(" ");
            String currentLine = words[0];
            for (int i = 1; i < words.length; i++) {
                if (fm.stringWidth(currentLine + words[i]) < QR_WIDTH - 32) {
                    currentLine += " " + words[i];
                } else {
                    tempStringList.add(currentLine);
                    currentLine = words[i];
                }
            }
            if (currentLine.trim().length() > 0) {
                tempStringList.add(currentLine);
            }
        } else {
            tempStringList.add(tempName);
        }
    }

    public static void main(String[] args) {
        try {
            //EXCEL FILE READER

            Workbook workbook = new XSSFWorkbook(new FileInputStream(new File("Test(New).xlsx")));
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            iterator.next();

            String name;
            int currentNo;
            int groupNo;
            String matricNo;
            String gender;
            String passNo;
            String email;
            String homeUni;
            String country;
            String offerLetterURL;
            String eValURL;
            DataFormatter formatter = new DataFormatter();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                //The number of fields should be equal to the number of column.
                currentNo = (int) cellIterator.next().getNumericCellValue();
                if(currentNo == 0){
                    break;
                }
                name = cellIterator.next().getStringCellValue();
//                if(name.contains("(")){
//                    name = name.substring(0, name.indexOf("("));
//                }
                passNo = formatter.formatCellValue(cellIterator.next());
                matricNo = cellIterator.next().getStringCellValue();
                email = cellIterator.next().getStringCellValue();
                gender = cellIterator.next().getStringCellValue();
                homeUni = cellIterator.next().getStringCellValue();
                cellIterator.next();
                country = cellIterator.next().getStringCellValue();
                cellIterator.next();
                cellIterator.next();
                cellIterator.next();
                cellIterator.next();
                cellIterator.next();
                groupNo = (int) cellIterator.next().getNumericCellValue();
                eValURL = cellIterator.next().getStringCellValue();
                offerLetterURL = cellIterator.next().getStringCellValue();
                if(offerLetterURL.isEmpty()) continue;
                File directory = new File("./" + FOLDERNAME + "/G" + groupNo);
                directory.mkdirs();

                //Generate QR Code

                generateQRCodeImage("https://docs.google.com/forms/d/e/1FAIpQLSdKK1qB4mzcusWtqF_7PTR6anqIPHsyosvM7fqrKN5UqpUohg/viewform?usp=pp_url" +
                                "&entry.2005620554=" + name +
                                "&entry.362091865=" + gender +
                                "&entry.1045781291=" + passNo +
                                "&entry.2120824388=" + matricNo +
                                "&entry.79707660=" + email +
                                "&entry.1065046570=" + groupNo +
                                "&entry.1166974658=" + country +
                                "&entry.839337160=" + homeUni +
                                "&entry.640261810=" + offerLetterURL +
                                "&entry.112646219=" + eValURL,
                        "./" + FOLDERNAME + "/G" + groupNo + "/" + name + ".png", name, matricNo, groupNo);
            }
        } catch (WriterException e) {
            System.out.println("Could not generate QR Code, WriterException :: " + e.getMessage());
        } catch (IOException e) {
            System.out.println("Could not generate QR Code, IOException :: " + e.getMessage());
        }
    }
}
