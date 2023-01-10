package va.easy_etudes_finder;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class Docx extends Fichier{
    public Docx(String docName, String docPath) {
        super();
        super.name = docName;
        super.path = docPath;
    }
    public String getText(){
        String text="";
        try{
            File f =  new File(path+name);
            FileInputStream fis = new FileInputStream(f.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            XWPFWordExtractor extracteur = new XWPFWordExtractor(document);
            text = extracteur.getText();//recupération du texte
            extracteur.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return text;

    }
}
