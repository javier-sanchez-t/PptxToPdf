/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.javic.pptxtopdf;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 *
 * @author acer
 */
public class Convert {
    // http://javapro.org/castellano/2017/07/25/convertir-archivo-powerpoint-pdf-java/
    public static void main(String[] args) {
        try {
            Convert.convertPPTToPDF("C:\\Users\\acer\\Desktop\\test.pptx", "C:\\Users\\acer\\Desktop\\p1.pdf", ".pptx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public static void convertPPTToPDF(String sourcepath, String destinationPath, String fileType) {
        double zoom = 2;
        AffineTransform at = new AffineTransform();
        at.setToScale(zoom, zoom);
        Document pdfDocument = new Document();
        byte[] barr = new byte[0];
        try (FileInputStream inputStream = new FileInputStream(sourcepath);
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, baos);
            PdfPTable table = new PdfPTable(1);
            pdfWriter.open();
            pdfDocument.open();
            Dimension pgsize = null;
            com.lowagie.text.Image slideImage = null;
            BufferedImage img = null;
            if (fileType.equalsIgnoreCase(".ppt")) {
                SlideShow ppt = new SlideShow(inputStream);
                pgsize = ppt.getPageSize();
                Slide slide[] = ppt.getSlides();
                pdfDocument.setPageSize(new com.lowagie.text.Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
                pdfWriter.open();
                pdfDocument.open();
                for (int i = 0; i < slide.length; i++) {
                    img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
                    Graphics2D graphics = img.createGraphics();
                    graphics.setTransform(at);
                    
                    graphics.setPaint(Color.white);
                    graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                    slide[i].draw(graphics);
                    graphics.getPaint();
                    slideImage = com.lowagie.text.Image.getInstance(img, null);
                    table.addCell(new PdfPCell(slideImage, true));
                }
            }
            if (fileType.equalsIgnoreCase(".pptx")) {
                XMLSlideShow ppt = new XMLSlideShow(inputStream);
                pgsize = ppt.getPageSize();
                //nuevo
                XSLFSlide slide1 = ppt.getSlides().get(0);
                XSLFNotes note = ppt.getNotesSlide(slide1);
                System.out.println("TEXTO==============>");
                /*for (XSLFTextShape shape : note.getPlaceholders()) {
                    System.out.println("texto: "+shape.getText());
                }*/
                //String nota=note.getPlaceholder(1).getText();
                
                java.util.List<org.apache.poi.xslf.usermodel.XSLFSlide> slide = ppt.getSlides();
                pdfDocument.setPageSize(new com.lowagie.text.Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
                pdfWriter.open();
                pdfDocument.open();
                for (XSLFSlide slid : slide) {
                    img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
                    Graphics2D graphics = img.createGraphics();
                    graphics.setTransform(at);
                    
                    graphics.setPaint(Color.white);
                    graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                    slid.draw(graphics);
                    graphics.getPaint();
                    slideImage = com.lowagie.text.Image.getInstance(img, null);
                    table.addCell(new PdfPCell(slideImage, true));
                    //table.addCell(nota);
                }
            }
            pdfDocument.add(table);
            pdfDocument.close();
            barr = baos.toByteArray();
            pdfWriter.close();
            System.out.println("Powerpoint file converted to PDF successfully");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        }
        
        try (FileOutputStream outputStream = new FileOutputStream(new File(destinationPath))) {
            outputStream.write(barr);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
