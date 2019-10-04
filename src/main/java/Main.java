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
    private final static String FOLDERNAME = "QR Output";

    private static void generateQRCodeImage(String text, String path, String name, String matricNo, int groupNumber)
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
        tempStringList.add("Group " + groupNumber);

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

            String name;
            int groupNumber;
            String matricNo;
            String gender;
            String ICPass;
            String email;
            String homeUniversity;
            String country;
            DataFormatter formatter = new DataFormatter();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                //The number of fields should be equal to the number of column.
                name = cellIterator.next().getStringCellValue();
                groupNumber = (int) cellIterator.next().getNumericCellValue();
                matricNo = cellIterator.next().getStringCellValue();
                ICPass = formatter.formatCellValue(cellIterator.next());
                gender = cellIterator.next().getStringCellValue();
                email = cellIterator.next().getStringCellValue();
                homeUniversity = cellIterator.next().getStringCellValue();
                country = cellIterator.next().getStringCellValue();

                File directory = new File("./" + FOLDERNAME + "/G" + groupNumber);
                directory.mkdirs();

                //Generate QR Code
                generateQRCodeImage("[google form url here]" +
                                "&entry.2005620554=" + name +
                                "&entry.1045781291=" + groupNumber +
                                "&entry.141207624=" + gender +
                                "&entry.1446259150=" + ICPass +
                                "&entry.688448863=" + matricNo +
                                "&entry.1166974658=" + country +
                                "&entry.1065046570=" + homeUniversity +
                                "&entry.444699550=" + email,
                        "./" + FOLDERNAME + "/G" + groupNumber + "/" + name + ".png", name, matricNo, groupNumber);
            }
        } catch (WriterException e) {
            System.out.println("Could not generate QR Code, WriterException :: " + e.getMessage());
        } catch (IOException e) {
            System.out.println("Could not generate QR Code, IOException :: " + e.getMessage());
        }
    }
}
