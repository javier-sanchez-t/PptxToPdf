/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.javic.pptxtopdf.util;

import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.Font;
import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.Image;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 *
 * @author acer
 */
public class Convert {

    // http://javapro.org/castellano/2017/07/25/convertir-archivo-powerpoint-pdf-java/
    public static void main(String[] args) {
        try {
            Convert.convertPPTToPDF("C:\\Users\\acer\\Desktop\\test.pptx", "C:\\Users\\acer\\Desktop\\p1.pdf", ".pptx", StaticConstants.PORTRAIT, 6);
        } catch (Exception ex) {
            Logger.getLogger(Convert.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void convertPPTToPDF(String sourcePathFile, String destinationPath, String fileType, String orientation, int fontSize) throws Exception {
        double zoom = 2;
        AffineTransform at = new AffineTransform();
        at.setToScale(zoom, zoom);

        //Final output file
        Document pdfDocument = new Document();

        FileInputStream inputStream = new FileInputStream(sourcePathFile);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, baos);
        //PdfPTable table = new PdfPTable(1);
        pdfWriter.open();
        pdfDocument.open();
        Dimension pgsize = null;
        Image slideImage = null;
        BufferedImage img = null;

        if (fileType.equalsIgnoreCase(".pptx")) {
            XMLSlideShow ppt = new XMLSlideShow(inputStream);
            pgsize = ppt.getPageSize();

            //pdfDocument.setPageSize(new com.lowagie.text.Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
            if (orientation.equals(StaticConstants.LANDSCAPE)) {
                pdfDocument.setPageSize(new Rectangle((float) 792, (float) 612));
            } else {
                pdfDocument.setPageSize(new Rectangle((float) 612, (float) 792));
            }
            pdfWriter.open();
            pdfDocument.open();

            for (int i = 0; i < ppt.getSlides().size(); i++) {
                XSLFSlide slid = ppt.getSlides().get(i);

                //Gets note 
                String note = getNoteFromSlide(ppt, slid);

                img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setTransform(at);

                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                slid.draw(graphics);
                graphics.getPaint();
                slideImage = Image.getInstance(img, null);

                //Adds slide image
                PdfPTable table = new PdfPTable(1);
                table.addCell(slideImage);
                pdfDocument.add(table);

                //Adds note
                Font font = new Font(FontFamily.TIMES_ROMAN, fontSize);
                pdfDocument.add(new Paragraph(note, font));
            }
        }

        pdfDocument.close();
        byte[] barr = baos.toByteArray();
        pdfWriter.close();
        System.out.println("Powerpoint file converted to PDF successfully");

        FileOutputStream outputStream = new FileOutputStream(new File(destinationPath));
        outputStream.write(barr);
    }

    /**
     * Get main note from specific slide
     *
     * @param ppt
     * @param slide
     * @return
     */
    public static String getNoteFromSlide(XMLSlideShow ppt, XSLFSlide slide) {
        XSLFNotes notes = ppt.getNotesSlide(slide);
        XSLFTextShape shape = notes.getPlaceholder(1);
        String note = "\n" + shape.getText() + "\n\n";
        if (!note.trim().equals("")) {
            note = "\n" + note + "\n";
        }
        return note;
    }
}
