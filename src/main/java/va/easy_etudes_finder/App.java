package va.easy_etudes_finder;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;


public class App {
    public static void main(String[] args) throws IOException {
        String excelFileName ="";
        String SheetName = "feuille1";
        String filePath = "/home/Jaffleman/Documents/banque-docx/";
        int columnNumber;
        try {
            InputStreamReader isr=new InputStreamReader(System.in);
            BufferedReader br=new BufferedReader(isr);
            System.out.print("Please provide the directory path: ");
            String fPath = br.readLine();
            if (fPath.length()>0) filePath = fPath;

            System.out.print("Please provide your Excel filename: ");
            String EFName = br.readLine();
            if (EFName.length()>0) excelFileName = EFName;
            else {
                System.out.println("You must provide a filename to continue.");
                return ;
            }
            System.out.print("Please enter sheet name: ");
            String SName = br.readLine();
            if (SName.length()>0) SheetName = SName;
            System.out.print("Please enter variable column : ");
            columnNumber = Converter.convertion(br.readLine());
            br.close();
            isr.close();
        } catch (Exception e) {throw e;}

        Xlsx excelFile = new Xlsx(filePath, excelFileName, SheetName, columnNumber);

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
        for (Docx docx : docTab) {
            String[] lineSplitText = docx.getText().split("\n");
            for(String variable:excelFile.getVariables()){ // Pour chaque variables
                final Pattern p = Pattern.compile(variable);
                String lineList = "";
                for (int i=0; i<lineSplitText.length; i++){ // Pour chaque lignes
                    String line = lineSplitText[i];
                    int i2= i+1;
                    Matcher m = p.matcher(line);
                    if (m.find()) lineList += " "+i2;
                    //else result = "\nfichier: "+item.getName()+" ------>\""+s+"\" non trouvé";
                }
                if(lineList.length() > 0) 
                    System.out.println("\nfichier: "+docx.name+" ------> \""+variable+"\" ------> ligne : ["+lineList+" ]");
            }
        }
        System.out.println("Terminé!");
    }
}
