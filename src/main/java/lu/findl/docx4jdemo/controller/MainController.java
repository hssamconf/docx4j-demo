package lu.findl.docx4jdemo.controller;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class MainController {

    @RequestMapping("/")
    public String index() {
        WordprocessingMLPackage wordMLPackage = null;
        try {
            wordMLPackage = WordprocessingMLPackage.createPackage();
            wordMLPackage.save(new java.io.File("C:\\Helloworld.docx"));

        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        return "Greetings from Spring Boot!";
    }
}
