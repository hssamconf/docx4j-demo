package lu.findl.docx4jdemo.helper;

//https://stackoverflow.com/questions/8354568/how-to-convert-a-ppt-file-into-pdf-file-in-java
// Sinon regarder sur Github : https://github.com/yeokm1/docs-to-pdf-converter/blob/master/docs-to-pdf-converter/src/com/yeokhengmeng/docstopdfconverter/PptxToPDFConverter.java

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

// import org.apache.poi.hslf.model.Slide;  // utiliser une version de poi-scratchpad <= 3.12
// import org.apache.poi.hslf.usermodel.SlideShow; // utiliser une version de poi-scratchpad <= 3.12
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class PPT2PDF {
	
	public static final String[] FILE_TYPES = new String[] { "ppt", "pptx" };

	@SuppressWarnings("resource")
	public static void convert(String inputPath, String outputPath) throws Exception {
	    FileInputStream inputStream = new FileInputStream(inputPath);
	    double zoom = 2;
	    AffineTransform at = new AffineTransform();
	    at.setToScale(zoom, zoom);
	    Document pdfDocument = new Document();
	    PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, new FileOutputStream(outputPath));
	    PdfPTable table = new PdfPTable(1);
	    pdfWriter.open();
	    pdfDocument.open();
	    Dimension pgsize = null;
	    Image slideImage = null;
	    BufferedImage img = null;
	    String fileType= "";
	    File pptfile = new File(inputPath);
		String fileName = pptfile.getName();
		if (fileName.toLowerCase().endsWith(PPT2PDF.FILE_TYPES[0])) {
			fileType = "ppt";
		} else {
			fileType = "pptx";
		}
	    System.out.println("fileType: "+fileType);
		    
	    if (fileType.equalsIgnoreCase("ppt")) {
	    	/*
	        SlideShow ppt = new SlideShow(inputStream);
	        pgsize = ppt.getPageSize();
	        Slide slide[] = ppt.getSlides();
	        pdfDocument.setPageSize(new Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
	        pdfWriter.open();
	        pdfDocument.open();
	        System.out.println(slide.length);
	        System.out.println(slide);
	        for (int i = 0; i < slide.length; i++) {
	            img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
	            Graphics2D graphics = img.createGraphics();
	            graphics.setTransform(at);

	            graphics.setPaint(Color.white);
	            graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
	            slide[i].draw(graphics);
	            graphics.getPaint();
	            slideImage = Image.getInstance(img, null);
	            table.addCell(new PdfPCell(slideImage, true));
	        }
	        */
	    	System.err.println("Error version poi-scratchpad");
	    }
	    else if (fileType.equalsIgnoreCase("pptx")) {
	        XMLSlideShow ppt = new XMLSlideShow(inputStream);
	        pgsize = ppt.getPageSize();
	        List<XSLFSlide> slide = ppt.getSlides();
	        pdfDocument.setPageSize(new Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
	        pdfWriter.open();
	        pdfDocument.open();
	        System.out.println(slide.size());
	        System.out.println(slide);
	        for (int i = 0; i < slide.size(); i++) {
	            img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
	            Graphics2D graphics = img.createGraphics();
	            graphics.setTransform(at);

	            graphics.setPaint(Color.white);
	            graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
	            slide.get(i).draw(graphics);
	            graphics.getPaint();
	            slideImage = Image.getInstance(img, null);
	            table.addCell(new PdfPCell(slideImage, true));
	        }
	    }
	    pdfDocument.add(table);
	    pdfDocument.close();
	    pdfWriter.close();
	    System.out.println("Powerpoint file converted to PDF successfully");
	}
}
