/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package tworzeniezestawow;

/**
 *
 * @author Grzegorz
 */
import com.spire.doc.*;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;
import com.spire.doc.fields.DocPicture;

import java.awt.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.util.concurrent.ThreadLocalRandom;

public class TworzenieZestawow {
    public static String[][] readTxtFile2OneDimArray(String csvFile) {

        //String csvFile = "c:\\!test_ta\\F1.csv";
        //BufferedReader br = null;
        String line = "";
        String cvsSplitBy = ";";
        String[][] datacsv = new String[10000][8];

        try {
            File fileDir = new File(csvFile);

            BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(fileDir), "UTF8"));


            //tablica tablic - kazdy element tablicy jest wektorem
            //o dwoch wspolrzednych: id, ta


            //br = new BufferedReader(new FileReader(csvFile));
            int counter = 0;
            while ((line = br.readLine()) != null) {

                // use comma as separator
                String[] splitLine = line.split(cvsSplitBy);
                System.out.println("splitLine = " + datacsv[counter]);
                datacsv[counter][0] = splitLine[0];
                datacsv[counter][1] = splitLine[1];
                datacsv[counter][2] = splitLine[2];
                datacsv[counter][3] = splitLine[3];
                datacsv[counter][4] = splitLine[4];
                datacsv[counter][5] = splitLine[5];
                datacsv[counter][6] = splitLine[6];
                datacsv[counter][7] = splitLine[7];
                System.out.println("id womi = " + datacsv[counter][0]+ "e-mail"+datacsv[counter][1]);
                counter++;

            }
            br.close();

        }
        catch (UnsupportedEncodingException e)
        {
            System.out.println(e.getMessage());
        }
        catch (IOException e)
        {
            System.out.println(e.getMessage());
        }
        catch (Exception e)
        {
            System.out.println(e.getMessage());
        }


        System.out.println("Przeczytany plik " + csvFile);
        return datacsv;
    }


    public static int losujLiczbeZPrzedzialu(int min, int max){
        return  ThreadLocalRandom.current().nextInt(min, max + 1);
    }
    public static void main(String[] args){
        //Read csv file: imie i nazwisko, adres e-mail, zad1 - wariant, zad2 - wariant, itd

        String[][] przeczytaneLinie = readTxtFile2OneDimArray("listaAlgebra.csv");
        for (int student = 1;student<=51;student++){
            if (przeczytaneLinie[student][1]==null)
                break;
            //Create a Document instance
            Document document = new Document();

            //Add a section
            Section section = document.addSection();

            //Add 3 paragraphs to the section
            Paragraph para1 = section.addParagraph();
            para1.appendText("Algebra liniowa - kolokwium nr 1");

            Paragraph para2 = section.addParagraph();
            //dodanie informacji id
            para2.appendText("Student: "+przeczytaneLinie[student][0]+", e-mail: "+przeczytaneLinie[student][1]);
            //
            //zadanie 1
            Paragraph para3 = section.addParagraph();
            Paragraph para4 = section.addParagraph();
            para3.appendText("Zadanie 1");
            para4.appendPicture(".\\Zadania\\AlgZad1v0"+przeczytaneLinie[student][2]+".bmp");

            //zadanie 2
            Paragraph para11 = section.addParagraph();
            Paragraph para12 = section.addParagraph();
            Paragraph para12pol = section.addParagraph();
            para11.appendText("Zadanie 2: Dane są macierze");
            para12.appendPicture(".\\Zadania\\AlgZad2v0"+przeczytaneLinie[student][3]+".bmp");
            //para12pol.appendPicture(".\\Zadania\\AlgZad5av00.bmp");

            //zadanie 3
            Paragraph para5 = section.addParagraph();
            Paragraph para6 = section.addParagraph();
            para5.appendText("Zadanie 3: Rozwiąż układy równań stosując poznane twierdzenia. Do jednego zastosuj wzory Cramera.");
            para6.appendText("(a) ");
            para6.appendPicture(".\\Zadania\\AlgZad3av0"+przeczytaneLinie[student][4]+".bmp");
            para6.appendText("(b) ");
            para6.appendPicture(".\\Zadania\\AlgZad3bv0"+przeczytaneLinie[student][5]+".bmp");


            //zadanie 5
            Paragraph para9 = section.addParagraph();
            Paragraph para10 = section.addParagraph();
            para9.appendText("Zadanie 4");
            para10.appendText("Teoria: Rachunek macierzowy i układy równań. Możliwa odpowiedź ustna na Teams. Termin do ustalenia mailowo.");


            //zadanie 5
            Paragraph para7 = section.addParagraph();
            Paragraph para8 = section.addParagraph();
            para7.appendText("Zadanie 5: ");
            para8.appendPicture(".\\Zadania\\AlgZad5v0"+przeczytaneLinie[student][6]+".bmp");
            //para12pol.appendPicture(".\\Zadania\\AlgZad5av00.bmp");


            //zadanie 6
            Paragraph para13 = section.addParagraph();
            Paragraph para14 = section.addParagraph();
            para13.appendText("Zadanie 6");
            para14.appendText("Teoria: Geometria analityczna. Możliwa odpowiedź ustna na Teams. Termin do ustalenia mailowo.");

            //Append an image from local file to the paragraph - bmp is OK
            //DocPicture picture = para2.appendPicture("c:\\Users\\Grzegorz\\Documents\\AlgZad1v01.bmp");
            //Set image width
            //picture.setWidth(300f);
            //Set image height
            //picture.setHeight(100f);

            //Set title style for paragraph 1
            ParagraphStyle style1 = new ParagraphStyle(document);
            style1.setName("titleStyle");
            style1.getCharacterFormat().setBold(true);
            style1.getCharacterFormat().setTextColor(Color.BLUE);
            style1.getCharacterFormat().setFontName("Arial");
            style1.getCharacterFormat().setFontSize(12f);
            document.getStyles().add(style1);
            para1.applyStyle("titleStyle");

            //Set style for paragraph 2 and 3
            ParagraphStyle style2 = new ParagraphStyle(document);
            style2.setName("paraStyle");
            style2.getCharacterFormat().setFontName("Arial");
            style2.getCharacterFormat().setFontSize(11f);
            document.getStyles().add(style2);
            para2.applyStyle("paraStyle");
            //para3.applyStyle("paraStyle");

            //Horizontally align paragraph 1 to center
            para1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

            //Set first-line indent for paragraph 2 and 3
            para2.getFormat().setFirstLineIndent(25f);
            //para3.getFormat().setFirstLineIndent(25f);

            //Set spaces after paragraph 1 and 2
            para1.getFormat().setAfterSpacing(15f);
            para2.getFormat().setAfterSpacing(10f);

            //Save the document - nazwa bedzie numerem albumu
            document.saveToFile(przeczytaneLinie[student][1]+".docx", FileFormat.Docx);
            //Save to file.
            document.saveToFile(przeczytaneLinie[student][1]+".pdf",FileFormat.PDF);
        }
    }
}
