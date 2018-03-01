package lu.findl.docx4jdemo.controller;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileNotFoundException;
import java.io.OutputStream;

@RestController
public class MainController {

    @RequestMapping("/docx")
    public String handlingDocx() {
        WordprocessingMLPackage wordMLPackage = null;
        try {
            // Exemple de creation d'un fichier docx et ajout d'un paragraphe hello world
            wordMLPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
            mdp.addParagraphOfText("hello world");
            String filename = System.getProperty("user.dir") + "\\OUT_hello.docx";
            Docx4J.save(wordMLPackage, new java.io.File(filename), Docx4J.FLAG_SAVE_ZIP_FILE);
            System.out.println("Saved " + filename);

            // Convertion du document générer en pdf
            String outputfilepath = System.getProperty("user.dir") + "/OUT_hello.pdf";
            OutputStream os = new java.io.FileOutputStream(outputfilepath);
            Docx4J.toPDF(wordMLPackage, os);
            System.out.println("Converted " + outputfilepath);

        } catch (Docx4JException | FileNotFoundException e) {
            e.printStackTrace();
        }

        return "Greetings from Spring Boot!";
    }

    @RequestMapping("/xlsx")
    public String handlingXlsx() {
        WordprocessingMLPackage wordMLPackage = null;


        return "Greetings from Spring Boot!";
    }
}
