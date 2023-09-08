
package com.example.worddemo;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Repository;


@Repository
public class Generate_Document {
    public static void Proba() {
        try {
            XWPFDocument document = new XWPFDocument();
            FileOutputStream out = new FileOutputStream(new File("C://Games//proba.docx"));
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("Szakdolgozat review");
            run.setBold(true);
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            File image = new File("C://Users//drig6//Demok//WordDemo//images//uni.jpg");
            FileInputStream imageData = new FileInputStream(image);
            int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
            String imageFileName = image.getName();
            int width = 450;
            int height = 400;
            run.addPicture(imageData, imageType, imageFileName,
                    Units.toEMU(width),
                    Units.toEMU(height));

            //create table
            XWPFTable table = document.createTable();
            //create first row
            XWPFTableRow tableRowOne = table.getRow(0);
            tableRowOne.getCell(0).setText("col one, row one");
            tableRowOne.addNewTableCell().setText("col two, row one");
            tableRowOne.addNewTableCell().setText("col three, row one");
            //create second row
            XWPFTableRow tableRowTwo = table.createRow();
            tableRowTwo.getCell(0).setText("col one, row two");
            tableRowTwo.getCell(1).setText("col two, row two");
            tableRowTwo.getCell(2).setText("col three, row two");
            //create third row
            XWPFTableRow tableRowThree = table.createRow();
            tableRowThree.getCell(0).setText("col one, row three");
            tableRowThree.getCell(1).setText("col two, row three");
            tableRowThree.getCell(2).setText("col three, row three");
            XWPFRun paragraphOneRunThree = paragraph.createRun();
            // paragraphOneRunThree.setStrike(true);
            paragraphOneRunThree.setFontSize(30);
            paragraphOneRunThree.setSubscript(VerticalAlign.SUBSCRIPT);
            paragraphOneRunThree.setText(" Font Styles");
            document.write(out);
            out.close();
            System.out.println("Word created succesfull");
        } catch (Exception e) {
            // TODO: handle exception
        }
    }
}
