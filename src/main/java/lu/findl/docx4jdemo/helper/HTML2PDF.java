package com.controller.sharepoint.convert;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Utilities;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.ElementList;
import com.itextpdf.tool.xml.XMLWorkerHelper;

import java.io.FileOutputStream;	
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang.StringEscapeUtils;

public class HTML2PDF {

	public static void convert(String inputPath, String outputPath) throws Exception {
		
		File file = new File(outputPath);
		file.getParentFile().mkdirs();
		try {
			new HTML2PDF().createPdf(inputPath,outputPath);

		} catch (DocumentException e) {
			e.printStackTrace();
		}
	}
	
	class TableHeader extends PdfPageEventHelper {
        String header;
        PdfTemplate total;
 
        public void setHeader(String header) {
            this.header = header;
        }
 
        public void onOpenDocument(PdfWriter writer, Document document) {
            total = writer.getDirectContent().createTemplate(30, 16);
        }
 
        public void onEndPage(PdfWriter writer, Document document) {
            PdfPTable table = new PdfPTable(3);
            DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm");
    		Date date = new Date();
    		
            try {
                table.setWidths(new int[]{24, 24, 2});
                table.setTotalWidth(527);
                table.setLockedWidth(true);
                table.getDefaultCell().setFixedHeight(20);
                table.getDefaultCell().setBorder(Rectangle.TOP);
                table.addCell(header);
                table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_RIGHT);
                table.addCell(String.format("%s - Page %d of", dateFormat.format(date), writer.getPageNumber()));
                PdfPCell cell = new PdfPCell(Image.getInstance(total));
                cell.setBorder(Rectangle.TOP);
                table.addCell(cell);
                table.writeSelectedRows(0, -1, 34, 60, writer.getDirectContent());
            }
            catch(DocumentException de) {
                throw new ExceptionConverter(de);
            }
        }
 
        /**
         * Fills out the total number of pages before the document is closed.
         * @see com.itextpdf.text.pdf.PdfPageEventHelper#onCloseDocument(
         *      com.itextpdf.text.pdf.PdfWriter, com.itextpdf.text.Document)
         */
        public void onCloseDocument(PdfWriter writer, Document document) {
            ColumnText.showTextAligned(total, Element.ALIGN_LEFT,
                    new Phrase(String.valueOf(writer.getPageNumber() - 1)),2, 2, 0);
        }
    }
	
	
	
	 /**
     * Creates a PDF
     * @param file
     * @throws IOException
     * @throws DocumentException
     */
    public void createPdf(String inputPath, String outputPath) throws IOException, DocumentException {
    	Document document = new Document(PageSize.A4, 40, 40, 40, 65);
        PdfWriter writer1 = PdfWriter.getInstance(document, new FileOutputStream(outputPath));
        TableHeader event = new TableHeader();
        writer1.setPageEvent(event);
        document.open();
          
        // Gestion du CSS
        String css = readCSS();
         
        // Gestion du HTML
        String html = Utilities.readFileToString(inputPath);
       
        // decode html (NIA ajout du 04/11/2016)
        html = StringEscapeUtils.unescapeHtml(html);
        ElementList list = XMLWorkerHelper.parseToElementList(html, css);
        for (Element e : list) {
            document.add(e);
        }

        /*
        if(mode.equals("print")){
        	PdfAction action = new PdfAction(PdfAction.PRINTDIALOG);
        	writer1.setOpenAction(action);
        }
        */
        document.close();
       
    }
 
    // Gestion du CSS
    private String readCSS() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int length;
        InputStream is = XMLWorkerHelper.class.getResourceAsStream("/default.css");
        while ((length = is.read(buffer)) != -1) {
            baos.write(buffer, 0, length);
        }
        return new String(baos.toByteArray());
    }
    
}
