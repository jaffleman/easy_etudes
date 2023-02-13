package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

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
        System.out.println("\nCreation of the Xlsx Class file...");
        Xlsx excelFile = new Xlsx(initDataRequested);
        report += "\nExcel File:";
        if (excelFile!=null) report += excelFile.getReport();
        File dir  = new File(initDataRequested.getPatn());
        File[] listeOfFiles = dir.listFiles();
        if (listeOfFiles == null){
            report += "\nPathFile error:";
            report += "\nSorry, no usable files in this directory!";
            return;
        }
        
        System.out.println("\nCreation of the docx Class files...");
        report += "\nDocx file :";
        List <Docx2> docxList = new ArrayList<>();
        for(File file : listeOfFiles){
            String fileName = file.getName();
            if(fileName.endsWith(".docx")) docxList.add(new Docx2(fileName, initDataRequested.getPatn()));
        }
        report += "\nNumber of docx founded: "+docxList.size();
        for (Docx2 docx : docxList) {
            report += docx.getReport();
        }

        System.out.println("\nStart seaching variables...");
        for (Docx2 docx2 : docxList) excelFile.workOnDcd(docx2);
        System.out.println("\nSaving results...");
        excelFile.saveDatatoSheet();
        Date date2 = new Date();
        long delta = date2.getTime()-date.getTime();
        int minutes = (int)(delta/60000);
        int seconds = (int)(delta/1000)-(minutes*60);
        System.out.println(report);
        System.out.println("\nTerminé en "+ minutes+" minutes "+seconds+" à: "+simpleDateFormat.format(date2));  
    }
}
