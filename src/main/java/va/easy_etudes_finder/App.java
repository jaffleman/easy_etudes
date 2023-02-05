package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

// import javax.swing.text.Segment;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;


public class App {
 
    public static void main(String[] args) throws IOException {
        String report = "";
        OperatingData initDataRequested = new OperatingData();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        Date date = new Date();
        report += "Débuté à: "+simpleDateFormat.format(date);
        //System.out.println(s.format(date));
        System.out.println("\nCreation of the Xlsx Class file...");
        Xlsx excelFile = new Xlsx(initDataRequested);
        report += "\nExcel File:";
        if (excelFile!=null) report += excelFile.getReport();
        //if(initData.getResultColumnNuber()==-1) excelFile = new Xlsx(filePath, excelFileName, SheetName, variableColumnNumber);
        //else excelFile;//filePath, excelFileName, SheetName, variableColumnNumber, resultColumnNumber);
        File dir  = new File(initDataRequested.getPatn());
        File[] listeOfFiles = dir.listFiles();
        if (listeOfFiles == null){
            report += "\nPathFile error:";
            report += "\nSorry, no usable files in this directory!";
            return;
        }
        
        System.out.println("\nCreation of the docx Class files...");
        report += "\nDocx file :";
        List <Docx> docxList = new ArrayList<>();
        for(File file : listeOfFiles){
            String fileName = file.getName();
            if(fileName.endsWith(".docx")) docxList.add(new Docx(fileName, initDataRequested.getPatn()));
        }
        report += "\nNumber of docx founded: "+docxList.size();
        for (Docx docx : docxList) {
            report += docx.getReport();
        }
        System.out.println("\nStart seaching variables...");
        List <String> excelSegStringList = excelFile.getSegmentList();
        List <Segments> excelSegElementList = excelFile.getSegList();
        for (Docx docx : docxList) { // pour chaque docx
            for (Segments docxSegElement : docx.getSegList()) { // pour chaque Segment de docx
                int index = excelSegStringList.indexOf(docxSegElement.getName()); // recupere le segment equivalent dans le fichier excel
                if (index!=-1){
                    for (SousSegment docxSousSegElement : docxSegElement.getSsList()) { // pour chaque SousSegment de docx
                        for (SousSegment excelSousSegElement : excelSegElementList.get(index).getSsList()) { //pour chaque SousSegment de xlsx
                            if(docxSousSegElement.getName().equals(excelSousSegElement.getName())){ // si le SousSegment de docx = le SousSegment de xlsx
                                for (String[] docxCodeTable : docxSousSegElement.getVarList()) { //pour chaque couple de variables (codeZone et codeBHA) de docx
                                    for (String[] excelCodeTable : excelSousSegElement.getVarList()) {//pour chaque couple de variables (codeZone et codeBHA) de docx
                                        if(docxCodeTable[0].equals(excelCodeTable[1])){ // si codeBHA de docx = codeBHA de xlsx
                                           // System.out.println( docxSegElement.getName()+" "+ docxSousSegElement.getName()+" "+ docxCodeTable[0]+" match "+excelCodeTable[1]);
                                           Pattern codeZone = Pattern.compile(docx.getshortName());
                                            Matcher codeZonMatcher = codeZone.matcher(excelCodeTable[2]);   
                                            if (!codeZonMatcher.find()) excelCodeTable[2]+=docx.getshortName()+";";
                                    }}
                                }
                            }
                        }
                    }
                }else{
                    System.out.println("Not found "+docxSegElement.getName());
                }
            }
        }
        for (Segments segments : excelSegElementList) {
            for (SousSegment sousSegments : segments.getSsList()) {
                for (String[] varList : sousSegments.getVarList()) {
                    System.out.println(varList[2]);
                    
                }
                
            }
        }
        System.out.println("\nSaving results...");
        excelFile.saveDatatoSheet();
        Date date2 = new Date();
        long delta = date2.getTime()-date.getTime();
        int minutes = (int)(delta/60000);
        int seconds = (int)(delta/1000)-(minutes*60);
        System.out.println(report);
        System.out.println("Terminé en "+ minutes+" minutes "+seconds+" à: "+simpleDateFormat.format(date2));  
    }
    private static int getFileNumber(File[] listeOfFiles) {
        int count = 0;
        for (File file : listeOfFiles){ 
            System.out.print(".");
            if(file.getName().endsWith(".docx")) count++;
        }
        return count;
    }
    private static boolean find(String elemToFind, String doc){
        final Pattern p = Pattern.compile(elemToFind);
        Matcher m = p.matcher(doc);
        return m.find();
    }
}
