package com.controller.sharepoint.convert;

import java.io.FileOutputStream;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class DOC2PDF {
	
	// TODO : Supprimer le copyright ( https://www.docx4java.org/forums/announces/docx4j-3-3-0-released-t2381.html )	
	
	public static void convert(String inputPath, String outputPath) throws Exception {
		WordprocessingMLPackage wordMLPackage;
		// 1) Load DOCX into WordprocessingMLPackage
		wordMLPackage = Docx4J.load(new java.io.File(inputPath));
		// 2) Convert WordprocessingMLPackage to Pdf
		Docx4J.toPDF(wordMLPackage, new FileOutputStream(outputPath));
	}
}
