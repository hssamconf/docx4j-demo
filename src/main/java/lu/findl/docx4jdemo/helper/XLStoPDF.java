package lu.findl.docx4jdemo.helper;

import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Iterator;

public class XLStoPDF {
    public static void convert(String inputPath, String outputPath) throws Exception {
        // Read the Excel file in binary format into FileInputStream
        InputStream input_document = new FileInputStream(inputPath);

        // Read workbook into XSSFWorkbook
        XSSFWorkbook workbook = new XSSFWorkbook(input_document);
        // Read worksheet into XSSFSheet
        XSSFSheet worksheet = workbook.getSheetAt(0);
        // To iterate over the rows
        Iterator<Row> rows = worksheet.iterator();
        // We will create output PDF document objects at this point
        Document pdf = new Document();
        PdfWriter.getInstance(pdf,	new FileOutputStream(outputPath));
        pdf.open();
        // Dynamic, Considering the above one as a header
        int noOfColumns = worksheet.getRow(0).getPhysicalNumberOfCells();
        PdfPTable xlsTable = new PdfPTable(noOfColumns);

        // To dynamically add new data to the table
        PdfPCell tableCell;

        while (rows.hasNext()) {
            Row row = rows.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println(row.getRowNum() + " | " + cell.getColumnIndex() + " | " + cell.getCellType()  + " | " + cell);

                // Identify CELL type
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        tableCell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                        xlsTable.addCell(tableCell);
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        String doubleValue = Double.toString(cell.getNumericCellValue());
                        tableCell = new PdfPCell(new Phrase(doubleValue));
                        xlsTable.addCell(tableCell);
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        String boolValue = Boolean.toString(cell.getBooleanCellValue());
                        tableCell = new PdfPCell(new Phrase(boolValue));
                        xlsTable.addCell(tableCell);
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        String formulaValue = Double.toString(cell.getNumericCellValue());
                        tableCell = new PdfPCell(new Phrase(formulaValue));
                        xlsTable.addCell(tableCell);
                        break;
                    case Cell.CELL_TYPE_BLANK:
                        tableCell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                        xlsTable.addCell(tableCell);
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        // Push the data from Excel to PDF Cell
                        tableCell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                        // feel free to move the code below to suit to your needs
                        xlsTable.addCell(tableCell);
                        break;
                }
            }
        }
        // Finally add the table to PDF document
        pdf.add(xlsTable);
        pdf.close();
        // we created our pdf file..
        input_document.close(); // close xls
        System.out.println("XLS " + inputPath + " to " + outputPath);
    }
}
