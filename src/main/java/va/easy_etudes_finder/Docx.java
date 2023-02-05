package va.easy_etudes_finder;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Docx extends Fichier{
    private String report = "";
    private List <String[]> text= new ArrayList<>();
    private List <Segments> segList = new ArrayList<>();
    public Docx(String docName, String docPath) {
        super();
        super.name = docName;
        super.path = docPath;
        
        String resultSegment = "";
        String sousSegment = "";
        
        String pathName = path+name;
        List<String[]> text = extractText(pathName);
        for (String[] strings : text) {
            List <String[]> listBHA = new ArrayList<>();
            List <String> result = new ArrayList<>();
            Pattern segment = Pattern.compile("^[0-9]{2}_[A-Z]+\\s");
            Matcher segMatcher = segment.matcher(strings[0]);
            while (segMatcher.find()){
                result.add(segMatcher.group().trim());
                resultSegment = segMatcher.group().trim();
            }
            Pattern subty = Pattern.compile("(subty|SUBTY|Subty)(\\s)*=( |\\s)*(\\W)?(\\s)*[A-Z_0-9]+");
            Matcher subMatcher = subty.matcher(strings[0]);
            while (subMatcher.find()){
                Pattern substract = Pattern.compile("(subty|SUBTY|Subty)(\\s)*=( |\\s)*(\\W)?(\\s)*");
                Matcher m = substract.matcher(subMatcher.group());
                result.add(m.replaceAll("").trim());
                sousSegment = m.replaceAll("").trim();
                // result.add(subMatcher.group());
            } 
            subMatcher = subty.matcher(strings[1]);
            while (subMatcher.find()){
                Pattern substract = Pattern.compile("(subty|SUBTY|Subty)(\\s)*=( |\\s)*(\\W)?(\\s)*");
                Matcher m = substract.matcher(subMatcher.group());
                result.add(m.replaceAll("").trim());
                sousSegment = m.replaceAll("").trim();
                // result.add(subMatcher.group());
            } 
            Pattern codeZone = Pattern.compile("(^|\\s| )[A-Z][A-Z_0-9]+(-| )[A-Z_0-9]+");
            Matcher codeZonMatcher = codeZone.matcher(strings[1]);
            while (codeZonMatcher.find()){
                result.add(codeZonMatcher.group().trim());
                listBHA.add(new String[]{codeZonMatcher.group().trim(),""}) ;
                
            }
            // System.out.println(strings[0]);
            // System.out.println("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
            for (String string : result) {
                    List <SousSegment> ssList = new ArrayList<>();
                    ssList.add(new SousSegment(sousSegment,listBHA));
                    segList.add(new Segments(resultSegment,ssList));
            }
        }
        System.out.println(this.name+" Terminé");
    }
    // private String extractText2(){
    //     String text="";
    //     try{
    //         System.out.flush();
    //         System.out.println(name);
    //         File f =  new File(path+name);
    //         FileInputStream fis = new FileInputStream(f.getAbsolutePath());
    //         XWPFDocument document = new XWPFDocument(fis);
    //         List<XWPFTable> tabDocs = document.getTables();
    //         for(XWPFTable table : tabDocs){
    //             if (table.getRow(0).getCell(0).getText().startsWith("Données")){ 
    //                 for(XWPFTableRow row:table.getRows()){
    //                     text += "\n"+row.getCell(2).getText();
    //         }}}
    //         document.close();fis.close();
    //     }catch(Exception e){}
    //     return text;
    // }

    private List<String[]> extractText(String pathName) {
        List <String[]> stringTab = new ArrayList<>();
        String txtColumn1="";
        String txtColumn2="";
        try{
            File f =  new File(pathName);
            FileInputStream fis = new FileInputStream(f.getAbsolutePath());
            ZipSecureFile.setMinInflateRatio(-1.0d);
            XWPFDocument document = new XWPFDocument(fis);
            List<XWPFTable> tabDocs = document.getTables();
            for(XWPFTable table : tabDocs){
                if (table.getRow(0).getCell(0).getText().startsWith("Données")){ 
                    for(XWPFTableRow row:table.getRows()){
                        txtColumn1 = row.getCell(2).getTextRecursively()+"\n";
                        txtColumn2 =row.getCell(3).getTextRecursively()+"\n";
                        stringTab.add(new String[]{txtColumn1,txtColumn2});
            }}}
            document.close();fis.close();
        }catch(Exception e){this.report +="\nError whith "+this.getshortName()+" while trying to read this file";}
        return stringTab;
    }
    
    public List <Segments> getSegList(){
        return this.segList;
    }

    public List<String[]> getText(){
        return this.text;
    }
    public String getshortName(){
        String[] splitName = this.name.split("_");
        String unParseName = splitName[splitName.length-1];
        unParseName = unParseName.substring(0, unParseName.length()-5);
        splitName = unParseName.split("-");
        return splitName[0];

    }
    public String getReport() {
        return report;
    }
}