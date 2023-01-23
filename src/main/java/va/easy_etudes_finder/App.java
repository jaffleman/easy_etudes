package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;


public class App {
 
    public static void main(String[] args) throws IOException {
        OperatingData initData = new OperatingData();
        SimpleDateFormat s = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        Date date = new Date();
        //System.out.println(s.format(date));
        System.out.println("\nCreation of the Xlsx Class file...");
        Xlsx excelFile = new Xlsx(initData);
        //if(initData.getResultColumnNuber()==-1) excelFile = new Xlsx(filePath, excelFileName, SheetName, variableColumnNumber);
        //else excelFile;//filePath, excelFileName, SheetName, variableColumnNumber, resultColumnNumber);
        File dir  = new File(initData.getPatn());
        File[] liste = dir.listFiles();
        if (liste == null){
            System.out.println("Sorry, no usable files in this directory!");
            return;
        }
        int count = 0;
        for (File file : liste){ 
            System.out.print(".");
            String fileName = file.getName();
            if(fileName.endsWith(".docx")) {
                count++;
            }
        }
        System.out.println("\nCreation of the docx Class files...");
        Docx[] docTab = new Docx[count];
        count = 0;
        for(File file : liste){
            String fileName = file.getName();
            if(fileName.endsWith(".docx")) {
                docTab[count] = new Docx(fileName, initData.getPatn());
                System.out.println(".");

                count++;
            }
        }
        System.out.println("\nStart seaching variables...");
        List <String[]> variablesDocStringList = new ArrayList<>();
        List<String[]> listOfV = excelFile.getCodeZoneList();
        for (String[] variables:listOfV) {// Pour chaque lignes de variables
            String codeZone = variables[0];
            String segment = variables[1];
            String sousSegment = variables[2];
            String docStringList = "";
            for(Docx docx : docTab){ //pour chaque docx
                String text = docx.getText();
                
                System.out.print(".");
                
                if (find(codeZone, text)){
                    if (
                        codeZone.equals("Debut")||
                        codeZone.equals("Fin")||
                        codeZone.equals("SUBTY")
                    ){
                        for (String splitTexString : text.split("\n")) {
                            
                            if (find(segment, splitTexString)) {
                                if (find(sousSegment, splitTexString)){
                                    if (find(codeZone, splitTexString)){
                                        docStringList += docx.getshortName()+";";
                                        break;
                                    }
                                }
                            }
                        }
                    }else docStringList += docx.getshortName()+";";
                }  
            }
            variablesDocStringList.add(new String[]{codeZone, docStringList});
        }
        System.out.println("\nSaving results...");
        excelFile.saveDatatoSheet(variablesDocStringList);
        Date date2 = new Date();
        long delta = date2.getTime()-date.getTime();
        int minutes = (int)(delta/60000);
        int seconds = (int)(delta/1000)-(minutes*60);
        System.out.println("Débuté à: "+s.format(date));
        System.out.println("Terminé en "+ minutes+" minutes "+seconds+" à: "+s.format(date2));
    }
    private static boolean find(String elemToFind, String doc){
        final Pattern p = Pattern.compile(elemToFind);
        Matcher m = p.matcher(doc);
        return m.find();
    }
}
