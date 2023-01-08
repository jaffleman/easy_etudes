package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.List;
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlsx extends Fichier {
    String sheetName;
    private int cellul;
    public Xlsx(String filePath, String fileName, String sheetName, int cellul) {
        super();
        super.name = fileName;
        super.path = filePath;  
        this.sheetName = sheetName;
        this.cellul = cellul;      
    }
    public List <String> getVariables(){
        List <String> listDeVariables = new ArrayList<String>() ;
        File f = new File(path+name);
        try {
            FileInputStream iss = new FileInputStream(f);
            System.out.print(".");
            FileInputStream fichier = new FileInputStream(f);
            XSSFWorkbook wb = new XSSFWorkbook(fichier);
            XSSFSheet feuille = wb.getSheet(sheetName);
            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            wb.close();iss.close();
            for (Row ligne : feuille) {//parcourir les lignes
                System.out.print("-");
                for (Cell cell : ligne) {//parcourir les colonnes
                    System.out.print("-");
                    //if( cell.getAddress()!= CellAddress.A1);
                    //évaluer le type de la cellule
                    System.out.print("-");
                    boolean isString = formulaEvaluator.evaluateInCell(cell).getCellType()==CellType.STRING?true:false;
                    boolean isRequireIndex = cell.getColumnIndex()== cellul? true:false;
                    if(isString && isRequireIndex) listDeVariables.add(cell.getStringCellValue());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return listDeVariables;
    }    
}
