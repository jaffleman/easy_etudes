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
        int i =0;
        List<String> listOfV = excelFile.getVariables();
        for (String variable:listOfV) {// Pour chaque variables
            i++;
            //System.out.println("\nsearching : "+variable+"    "+i+"/"+listOfV.size());
            String docStringList = "";
            for(Docx docx : docTab){ //pour chaque docx
                System.out.print(".");
                final Pattern p = Pattern.compile(variable);
                Matcher m = p.matcher(docx.getText());
                if (m.find()) docStringList += docx.getshortName()+";";
            }
            variablesDocStringList.add(new String[]{variable, docStringList});
        }
        System.out.println("\nSaving results...");
        excelFile.saveDatatoSheet(variablesDocStringList);
        Date date2 = new Date();
        long delta = date2.getTime()-date.getTime();
        
        System.out.println("Débuté à: "+s.format(date));
        System.out.println("Terminé en "+ delta/6000+"minutes à: "+s.format(date2));
    }
}
