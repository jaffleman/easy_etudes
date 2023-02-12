package va.easy_etudes_finder;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Docx2 {
    private String status="ok";
    private String name, path;
    private String report = "";
    Boolean foundedData = false;
    private String text = "";
    private List <String[]> textLineList= new ArrayList<>();
    private List <Segments> segList = new ArrayList<>();
    private List<String> docSegList = new ArrayList<>();
    static List<String> segListIndex = new ArrayList<>();
    static List<String> shortNameListIndex = new ArrayList<>();
    public Docx2(String docName, String docPath) throws IOException {
        System.out.print(".");
        this.name = docName;
        this.path = docPath;
        // List <String> celluleExtractData = new ArrayList<>();
        shortNameListIndex.add(this.getshortName());
        if(this.getshortName().equals("GRADE")){
            System.out.println("");
        }
        List <String[]> docxExtractData = new ArrayList<>();
        String pathName = path+name;
        List<String[]> textLineList = extractText(pathName);
        String previousCS="";
        String previousSS="";
        for (String[] strings : textLineList) {
            String codeSegment = "";
            String sousSegment = "";
            String results = "";
            Pattern segment = Pattern.compile("[0-9]{2}_[A-Z]+\\s");
            Matcher segMatcher = segment.matcher(strings[0]);
            while (segMatcher.find()){
                codeSegment = segMatcher.group();
            }
            Pattern subty = Pattern.compile("(subty|SUBTY|Subty)( |\\s)*=( |\\s)*(\\W)?( |\\s)*[A-Z_0-9]+");
            Matcher subMatcher = subty.matcher(strings[0]);
            while (subMatcher.find()){
                Pattern substract = Pattern.compile("(subty|SUBTY|Subty)( |\\s)*=( |\\s)*(\\W)?( |\\s)*");
                Matcher m = substract.matcher(subMatcher.group());
                sousSegment = m.replaceAll("");
                String subResult = m.replaceAll("");
                Pattern number = Pattern.compile("0[0-9]");
                Matcher cpresult = number.matcher(subResult);
                if (cpresult.find()) {
                    substract = Pattern.compile("0");
                    m = substract.matcher(cpresult.group());
                    sousSegment = m.replaceAll("");
                }
            }
            if (sousSegment.equals("")){
                subty = Pattern.compile("(subty|SUBTY|Subty)( |\\s)*=( |\\s)*(\\W)?( |\\s)*[A-Z_0-9]+");
                subMatcher = subty.matcher(strings[1]);
                while (subMatcher.find()){
                    Pattern substract = Pattern.compile("(subty|SUBTY|Subty)( |\\s)*=( |\\s)*(\\W)?( |\\s)*");
                    Matcher m = substract.matcher(subMatcher.group());
                    sousSegment = m.replaceAll("");
                    String subResult = m.replaceAll("");
                    Pattern number = Pattern.compile("0[0-9]");
                    Matcher cpresult = number.matcher(subResult);
                    if (cpresult.find()) {
                        substract = Pattern.compile("0");
                        m = substract.matcher(cpresult.group());
                        sousSegment = m.replaceAll("");
                    }
                }
            }
            Pattern debut = Pattern.compile("(début|Debut|Début|DEBUT) (de )?segment");
            Matcher debutMatcher = debut.matcher(strings[0]);
            while (debutMatcher.find()){
                results+="Debut;";
            }
            Pattern fin = Pattern.compile("(Fin|fin|FIN) (de )?segment");
            Matcher finMatcher = fin.matcher(strings[0]);
            while (finMatcher.find()){
                results+="Fin;";
            }
            Pattern sub = Pattern.compile("(subty|SUBTY|Subty)");
            Matcher subMatcher2 = sub.matcher(strings[0]);
            while (subMatcher2.find()){
                results+="SUBTY;";
            }
            Pattern codeZone = Pattern.compile("(A_|O_|P_|F_)[A-Z_0-9][A-Z_0-9][A-Z_0-9]_[A-Z_0-9]+");
            Matcher codeZonMatcher = codeZone.matcher(strings[0]);
            while (codeZonMatcher.find()){
                results+=codeZonMatcher.group()+";";
            }
            if(codeSegment.equals("") && sousSegment.equals("") && results.equals("")){
            }else if (codeSegment.equals("") && sousSegment.equals("") && !results.equals("")) {
                String[] celluleDataGroup =new String[3];
                celluleDataGroup[0]=previousCS;
                celluleDataGroup[1]=previousSS;
                celluleDataGroup[2]=results;
                docxExtractData.add(celluleDataGroup);
            }else if (codeSegment.equals("") && !sousSegment.equals("") && !results.equals("")) {
                String[] celluleDataGroup =new String[3];
                celluleDataGroup[0]=previousCS;
                celluleDataGroup[1]=sousSegment.trim();
                celluleDataGroup[2]=results;
                docxExtractData.add(celluleDataGroup);
            }else{
                String[] celluleDataGroup =new String[3];
                celluleDataGroup[0]=codeSegment.trim();
                celluleDataGroup[1]=sousSegment.trim();
                celluleDataGroup[2]=results;
                docxExtractData.add(celluleDataGroup);
                previousCS = codeSegment.trim();
                previousSS = sousSegment.trim();
            }
        }
        int prvIndex1 = 0, prvIndex2 = 0;
        for (String[] celluleData : docxExtractData){
            int segIndex = 0;
            int sousSegIndex = 0;
            Boolean segFound = false;
            Boolean sousSegFound = false;
            if(!celluleData[0].equals("")){
                for(int i=0;i<segList.size();i++){
                    Segments segment = segList.get(i);
                    if (segment.getName().equals(celluleData[0])){
                        segIndex = i;
                        segFound = true;
                        break;
                    }
                }
            }
            if (segList.size()>0){
                Segments segment = segList.get(segIndex);
                List <SousSegment> sousSegList = segment.getSsList();
                for(int i=0; i<sousSegList.size();i++){
                    SousSegment sousSegment = sousSegList.get(i);
                    if(sousSegment.getName().equals(celluleData[1])){
                        sousSegIndex = i;
                        sousSegFound = true;
                        break;
                    }
                }
            }
            
            if (segFound) {
                if (sousSegFound) {
                    List <String> stringTabString = segList.get(segIndex).getSsList().get(sousSegIndex).getDocxVarList();
                    for ( String elem : celluleData[2].split(";")) {
                        if (!stringTabString.contains(elem))
                        segList.get(segIndex).getSsList().get(sousSegIndex).getDocxVarList().add(elem);
                    }                    
                }else{
                    List <String> stringTabString = new ArrayList<>();
                    for ( String elem : celluleData[2].split(";")) {
                        if (!stringTabString.contains(elem))
                        stringTabString.add(elem);
                    }
                    segList.get(segIndex).getSsList().add(new SousSegment(celluleData[1], stringTabString, 0));
                }                
            }else{

                if (celluleData[0].equals("")) {
                    List <String> stringTabString = segList.get(prvIndex1).getSsList().get(prvIndex2).getDocxVarList();
                    for ( String elem : celluleData[2].split(";")) {
                        if (!stringTabString.contains(elem))
                        segList.get(prvIndex1).getSsList().get(prvIndex2).getDocxVarList().add(elem);
                    }
                } else {
                    foundedData=true;
                    docSegList.add(celluleData[0]);
                    List <String> stringTabString = new ArrayList<>();
                    for ( String elem : celluleData[2].split(";")) {
                        if (!stringTabString.contains(elem))
                        stringTabString.add(elem);
                    }
                    List <SousSegment> ssegList = new ArrayList<>();
                    ssegList.add(new SousSegment(celluleData[1], stringTabString,0));
                    this.segList.add(new Segments(celluleData[0], ssegList));
                    if (!segListIndex.contains(celluleData[0])) segListIndex.add(celluleData[0]);
                }
            }
            prvIndex1 = segIndex;
            prvIndex2 = sousSegIndex;
        }
        if (foundedData==null) {
            this.report +="\n"+getshortName()+" No such data founded! please check manually";
            this.status = "ERROR";
        }
    }

    private List<String[]> extractText(String pathName) {
        List <String[]> stringTab = new ArrayList<>();
        String txtColumn1="";
        String txtColumn2="";
        File f =  new File(pathName);
        FileInputStream fis = null;
        XWPFDocument document = null;
        try{ fis = new FileInputStream(f.getAbsolutePath());}
        catch(Exception e){
            this.status="ERROR";
            this.report +="\n"+this.getshortName()+": Error while opening the input Sream.";
        }
        if (fis != null) {
            try {
                ZipSecureFile.setMinInflateRatio(-1.0d);
                document = new XWPFDocument(fis);
            } catch (Exception e) {
                this.status="ERROR";
                this.report +="\n"+this.getshortName()+": Error while reading the file.";
        }}
        if (document != null) {
            List<XWPFTable> tabDocs = document.getTables();
            for(XWPFTable table : tabDocs){
                if (table.getRow(0).getCell(0).getText().startsWith("Données")){ 
                    for(XWPFTableRow row:table.getRows()){
                        try {
                            txtColumn1 = row.getCell(2).getTextRecursively()+"\n";
                            txtColumn2 =row.getCell(3).getTextRecursively()+"\n";
                            text += txtColumn1 + txtColumn2;
                            stringTab.add(new String[]{txtColumn1,txtColumn2});
                        } catch (Exception e) {}
            }}}
            try {document.close();fis.close();} 
            catch (IOException e) {
                this.report +="\n"+this.getshortName()+": Error while closing the file or the stream.";
        }}
        return stringTab;
    }
    
    public List <Segments> getSegList(){
        return this.segList;
    }

    public List<String[]> getText(){
        return this.textLineList;
    }
    public String getshortName(){
        String[] splitName = this.name.split("_");
        String unParseName = splitName[splitName.length-1];
        unParseName = unParseName.substring(0, unParseName.length()-5);
        splitName = unParseName.split("-| ");
        if (splitName[0].equals("2020")) report += "\nIl est la ton fichier 2020 :"+ this.name; 
        return splitName[0];

    }
    public String getExtractText(){
        return this.text;
    }
    public String getReport() {
        return report;
    }

    public String getStatus() {
        return this.status;
    }
}
