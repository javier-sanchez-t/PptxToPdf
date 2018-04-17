package com.javic.pptxtopdf.util;

import com.itextpdf.text.BadElementException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.FontFactory;
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
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * @version 1.0
 * @since 14/04/2018
 * @author Javier Sánchez, Upwork
 */
public class Convert {

    public void convertPPTToPDF(String sourcePathFile, String destinationPath, String fileType, String orientation, String fontName, int fontSize) throws Exception {
        Runtime garbage = Runtime.getRuntime();

        double zoom = 1;
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

            //Defines page orientation 
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
                Paint paint = graphics.getPaint();
                slideImage = Image.getInstance(img, null);
                garbage.gc();

                //Adds slide image
                PdfPTable table = new PdfPTable(1);
                table.addCell(slideImage);
                pdfDocument.add(table);

                //Adds note
                Font font = FontFactory.getFont(fontName);
                font.setSize(fontSize);
                pdfDocument.add(new Paragraph(note, font));

                pdfDocument.newPage();
                slid = null;
                note = null;
                font = null;
                table = null;
                slideImage = null;
                graphics = null;
                paint = null;
                img = null;
                slid = null;
                garbage.gc();
            }

            pdfDocument.close();
            ppt = null;
            pdfDocument = null;
            garbage.gc();
        }

        byte[] barr = baos.toByteArray();
        pdfWriter.close();
        System.out.println("Powerpoint file converted to PDF successfully");

        FileOutputStream outputStream = new FileOutputStream(new File(destinationPath));
        outputStream.write(barr);

        baos = null;
        barr = null;
        outputStream = null;
        inputStream = null;
        pdfWriter = null;
        garbage.gc();
    }

    public static void convertPPTToPDF_oneFile(java.util.List<File> pptxFiles, String destinationPath, String orientation, String fontName, int fontSize) throws Exception {

        //Final output file
        Document pdfDocument = new Document();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        //writer
        PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, baos);
        pdfWriter.open();
        pdfDocument.open();

        //Creates one PDF file from many PPTX
        for (File file : pptxFiles) {
            createPagesFromSlides(pdfWriter, pdfDocument, file.getAbsolutePath(), orientation, fontName, fontSize);
        }

        pdfDocument.close();
        byte[] barr = baos.toByteArray();
        pdfWriter.close();
        System.out.println("Powerpoint file converted to PDF successfully");

        FileOutputStream outputStream = new FileOutputStream(new File(destinationPath));
        outputStream.write(barr);
    }

    public static void createPagesFromSlides(
            PdfWriter pdfWriter,
            Document pdfDocument,
            String sourcePathFile,
            String orientation,
            String fontName,
            int fontSize)
            throws IOException, BadElementException, DocumentException {

        FileInputStream inputStream = new FileInputStream(sourcePathFile);

        double zoom = 2;
        AffineTransform at = new AffineTransform();
        at.setToScale(zoom, zoom);

        //PdfPTable table = new PdfPTable(1);
        Dimension pgsize = null;
        Image slideImage = null;
        BufferedImage img = null;

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
            Font font = FontFactory.getFont(fontName);
            font.setSize(fontSize);
            pdfDocument.add(new Paragraph(note, font));
        }
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
