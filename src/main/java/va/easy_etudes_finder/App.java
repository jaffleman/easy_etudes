package va.easy_etudes_finder;


import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) throws IOException {
        String excelFileName ="";
        String SheetName = "feuille1";
        String filePath = "/home/Jaffleman/Documents/banque2";
        int cNumber;
        try {
            InputStreamReader isr=new InputStreamReader(System.in);
            BufferedReader br=new BufferedReader(isr);
            System.out.print("Please provide the directory path: ");
            String fPath = br.readLine();
            if (fPath.length()>0) filePath = fPath;

            // Class Xlsx _____________________________________________________
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
            System.out.print("Please enter variable column number: ");
            cNumber = Integer.parseInt(br.readLine()) ;
            br.close();
            isr.close();
        } catch (Exception e) {throw e;}
        
        List <String> listDeVariables = new ArrayList<String>() ;
        File dir  = new File(filePath);
        File[] liste = dir.listFiles();
        if (liste == null){
            System.out.println("Sorry, no usable files in this directory!");
            return;
        }
        boolean isFound = false;
        for(File item : liste){

            System.out.print("-");
            String itemName = item.getName();
            if(itemName.equals(excelFileName) ){
                isFound=true;
                //System.out.format("Nom du fichier: %s%n", item.getName()); 
                File f=  new File(item.getAbsolutePath());
                FileInputStream iss = new FileInputStream(f);
                   System.out.print(".");
                    try {
                        FileInputStream fichier = new FileInputStream(f);
                        XSSFWorkbook wb = new XSSFWorkbook(fichier);
                        XSSFSheet feuille = wb.getSheet(SheetName);
                        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
                        wb.close();
                        for (Row ligne : feuille) {//parcourir les lignes
                            System.out.print("-");
                            for (Cell cell : ligne) {//parcourir les colonnes
                                System.out.print("-");
                                //if( cell.getAddress()!= CellAddress.A1);
                                //évaluer le type de la cellule
                                System.out.print("-");
                                boolean isString = formulaEvaluator.evaluateInCell(cell).getCellType()==CellType.STRING?true:false;
                                boolean isRequireIndex = cell.getColumnIndex()== cNumber? true:false;
                                if(isString && isRequireIndex) listDeVariables.add(cell.getStringCellValue());
                            }
                        }
                    } catch (Exception e) {
                        throw e;
                    }
                    finally {
                        if (iss != null) iss.close();
                    }
                
            }
        }
        if(!isFound) {
            System.out.println("\n"+excelFileName+" is not found in the provide directory.");
            return;
        }
        for(File item : liste){ //Pour chaque fichier docx
            //System.out.println(item.getName());
            if(item.getName().endsWith(".docx")){
                System.out.println("\nSearching docx files "+item.getName());
                try{
                    File f =  new File(item.getAbsolutePath());
                    FileInputStream fis = new FileInputStream(f.getAbsolutePath());
                    XWPFDocument document = new XWPFDocument(fis);
                    XWPFWordExtractor extracteur = new XWPFWordExtractor(document);
                    //String result = "Null";
                    String text = extracteur.getText();//recupération du texte
                    extracteur.close();
                    String[] lineSplitText = text.split("\n");
                    for(String variable:listDeVariables){ // Pour chaque variables
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
                            System.out.println("\nfichier: "+item.getName()+" ------> \""+variable+"\" ------> ligne : ["+lineList+" ]");
                    }
                    fis.close();
                }
                catch (Exception e) {
                                e.printStackTrace();
                                }

            }
            else{} //System.out.println("NoMatch");
        }
        System.out.println("Terminé!");
    }
}
