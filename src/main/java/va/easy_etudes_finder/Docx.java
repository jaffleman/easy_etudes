package va.easy_etudes_finder;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Docx extends Fichier{
    private String text;
    public Docx(String docName, String docPath) {
        super();
        super.name = docName;
        super.path = docPath;
        this.text = extractText();
        System.out.print(".");
    }
    private String extractText(){
        String text="";
        try{
            System.out.println(name);
            File f =  new File(path+name);
            FileInputStream fis = new FileInputStream(f.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            List<XWPFTable> tabDocs = document.getTables();
            for(XWPFTable table : tabDocs){
                if (table.getRow(0).getCell(0).getText().startsWith("Données")){ 
                    for(XWPFTableRow row:table.getRows()){
                        text += "\n"+row.getCell(2).getText();
            }}}
            document.close();fis.close();
        }catch(Exception e){}
        //System.out.println(text);
        return text;
    }
    public String getText(){
        return this.text;
    }
    public String getshortName(){
        String[] splitName = this.name.split("_");
        String unParseName = splitName[splitName.length-1];
        unParseName = unParseName.substring(0, unParseName.length()-5);
        splitName = unParseName.split("-");
        return splitName[0];

    }
}
