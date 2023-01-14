package va.easy_etudes_finder;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;


public class App {
    enum SameResultSheet {OUI, NON};
    public static void main(String[] args) throws IOException {
        String excelFileName ="";
        String SheetName = "feuille1";
        String filePath = "/home/Jaffleman/Documents/banque-docx/";
        int variableColumnNumber, resultColumnNumber=-1;
        InputStreamReader isr=new InputStreamReader(System.in);
        BufferedReader br=new BufferedReader(isr);
        System.out.print("Please provide the directory path: ");
        String fPath = br.readLine();
        if (fPath.length()>0) filePath = fPath;

        System.out.print("Please provide your Excel filename: ");
        String EFName = br.readLine();
        while(EFName.length()<1) {
            System.out.println("You must provide a filename to continue.");
            EFName = br.readLine();
        }
        excelFileName = EFName;
        System.out.print("Please enter sheet name (if different from: 'feuille1'): ");
        String SName = br.readLine();
        if (SName.length()>0) SheetName = SName;
        System.out.print("Please enter variable column : ");
        variableColumnNumber= Converter.convertion(br.readLine());
        System.out.print("Same sheet for results (yes/no)?");
        boolean sameSheet = br.readLine().equals("no")?false:true;
        if(sameSheet) {
            System.out.print("Please enter result column : ");
            resultColumnNumber = Converter.convertion(br.readLine()); 
        }
        br.close();isr.close();
        Xlsx excelFile;
        if(resultColumnNumber==-1) excelFile = new Xlsx(filePath, excelFileName, SheetName, variableColumnNumber);
        else excelFile = new Xlsx(filePath, excelFileName, SheetName, variableColumnNumber, resultColumnNumber);
        File dir  = new File(filePath);
        File[] liste = dir.listFiles();
        if (liste == null){
            System.out.println("Sorry, no usable files in this directory!");
            return;
        }
        int count = 0;
        for (File file : liste){ 
            String fileName = file.getName();
            if(fileName.endsWith(".docx")) {
                count++;
            }
        }
        Docx[] docTab = new Docx[count];
        count = 0;
        for(File file : liste){
            String fileName = file.getName();
            if(fileName.endsWith(".docx")) {
                docTab[count] = new Docx(fileName, filePath);
                count++;
            }
        }
        for (String variable:excelFile.getVariables()) {// Pour chaque variables
            String docStringList = "";
            for(Docx docx : docTab){ //pour chaque docx
                final Pattern p = Pattern.compile(variable);
                Matcher m = p.matcher(docx.getText());
                if (m.find()) docStringList += " "+docx.getshortName()+";";
            }
            excelFile.saveDatatoSheet(variable, docStringList);
        }
        System.out.println("Terminé!");
    }
}
