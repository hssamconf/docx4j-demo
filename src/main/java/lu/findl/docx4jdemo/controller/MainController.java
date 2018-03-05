package lu.findl.docx4jdemo.controller;

import lu.findl.docx4jdemo.helper.DOC2PDF;
import lu.findl.docx4jdemo.helper.HTML2PDF;
import lu.findl.docx4jdemo.helper.XLS2HTML;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.*;

import java.io.File;

@RestController
public class MainController {

    private static void addContent(WorksheetPart sheet) {

        // Minimal content already present
        SheetData sheetData = sheet.getJaxbElement().getSheetData();

        // Now add
        Row row = Context.getsmlObjectFactory().createRow();
        //row.setR((long) 1);  // optional

        Cell cell = Context.getsmlObjectFactory().createCell();
        cell.setV("1234");
        cell.setR("A1");  // Apple Numbers needs this, or cell content won't show
        row.getC().add(cell);

        Cell cell2 = createCell("hello world!");
        row.getC().add(cell2);
        cell2.setR("B1"); // be careful the numeral matches the row correctly, or Excel will complain

        sheetData.getRow().add(row);
    }

    private static Cell createCell(String content) {

        Cell cell = Context.getsmlObjectFactory().createCell();

        CTXstringWhitespace ctx = Context.getsmlObjectFactory().createCTXstringWhitespace();
        ctx.setValue(content);

        CTRst ctrst = new CTRst();
        ctrst.setT(ctx);

        cell.setT(STCellType.INLINE_STR);
        cell.setIs(ctrst); // add ctrst as inline string

        return cell;

    }

    @RequestMapping(value = {"/docx", "/doc"})
    public String handlingDocx() {
        WordprocessingMLPackage wordMLPackage = null;
        try {
            // Exemple de creation d'un fichier docx et ajout d'un paragraphe hello world
            wordMLPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
            mdp.addParagraphOfText("hello world");
            String filename = System.getProperty("user.dir") + "\\OUT_hello_doc.docx";
            Docx4J.save(wordMLPackage, new java.io.File(filename), Docx4J.FLAG_SAVE_ZIP_FILE);
            System.out.println("DOCX Saved .." + filename);

            // Convertion du document générer en pdf
            String outputfilepath = System.getProperty("user.dir") + "\\OUT_hello_doc.pdf";
            DOC2PDF.convert(filename, outputfilepath);
            System.out.println("DOCX Converted .." + outputfilepath);

        } catch (Exception e) {
            e.printStackTrace();
        }

        return "Greetings from Spring Boot!";
    }

    @RequestMapping(value = {"/xlsx", "/xls"})
    public String handlingXlsx() {
        try {
            // Exemple de creation d'un fichier xlsx
            String outputfilepath = System.getProperty("user.dir") + "\\OUT_hello_xsl.xlsx";
            SpreadsheetMLPackage pkg = SpreadsheetMLPackage.createPackage();
            WorksheetPart sheet = pkg.createWorksheetPart(new PartName("/xl/worksheets/sheet1.xml"), "Sheet1", 1);
            //ajout d'un contenu dans le fichier xlsx (deux cellules 1234 et Hello world)
            addContent(sheet);
            pkg.save(new File(outputfilepath));
            System.out.println("XSLX Saved .." + outputfilepath);

            //Converting xlsx to pdf using itext
            String outputhtmlfilepath = System.getProperty("user.dir") + "\\OUT_hello_xsl.html";
            String outputpdffilepath = System.getProperty("user.dir") + "\\OUT_hello_xsl.pdf";

            // TODO : Create Helper XLS2HTML2PDF direct.
            //XLStoPDF.convert(outputfilepath, outputpdffilepath);
            XLS2HTML.convert(outputfilepath, outputhtmlfilepath);
            HTML2PDF.convert(outputhtmlfilepath, outputpdffilepath);

        } catch (Exception e) {
            e.printStackTrace();
        }

        return "Greetings from Spring Boot!";
    }
}
