package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlsx extends Fichier {
    private String result_sheet = "result_sheet";
    private String sheetName;
    private int variableColumn;
    private int resultColumn;
    private XSSFWorkbook wb;
    private boolean newSheetForResults=false;

    public Xlsx(String filePath, String fileName, String sheetName, int variableColumn) {
        // new sheet
        super();
        this.newSheetForResults=true;
        super.name = fileName;
        super.path = filePath;  
        this.sheetName = sheetName;
        this.variableColumn = variableColumn; 
        List<String> varables = this.getVariables();
        XSSFSheet feuille = wb.getSheet(result_sheet);
        if (feuille==null){
            feuille = wb.createSheet(result_sheet);
            Row row0 = feuille.createRow(0);
            // EmpNo
            Cell cell = row0.createCell(0, CellType.STRING);
            cell.setCellValue("SEARCH RESULTS:");
            feuille.addMergedRegion(new CellRangeAddress(
                0, //first row (0-based)
                0, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
            ));
            // cell.setCellStyle(style);
            // EmpName
            Row row1 = feuille.createRow(1);
            cell = row1.createCell(0, CellType.STRING);
            cell.setCellValue("Variables list:");
            // cell.setCellStyle(style);
            cell = row1.createCell(3, CellType.STRING);
            cell.setCellValue("found Etudes:");
            // cell.setCellStyle(style);
            for (int i = 2; i <= varables.size()+1; i++ ) {
                row1 = feuille.createRow(i);
                cell = row1.createCell(0, CellType.STRING);
                cell.setCellValue(varables.get(i-2));
            }
            writeFlux();
        } 
    }

    public Xlsx(String filePath, String fileName, String sheetName, int variableColumn, int rColumn) {
        // same sheet
        super();
        super.name = fileName;
        super.path = filePath;  
        this.sheetName = sheetName;
        this.variableColumn = variableColumn; 
        this.resultColumn = rColumn;
        
        // initReadFlux();
        // XSSFSheet feuille = wb.getSheet(sheetName);
        // Cell cell = feuille.getRow(1).createCell(resultColumn);
        // cell.setCellValue("found etudes: ");
        // writeFlux(); 
    }

    /**
     * Initialise le flux de lecture
     * pour la recupération des données du fichier
     */
    private void initReadFlux(){
        File f = new File(path+name);
        try {
            // FileInputStream iss = new FileInputStream(f);
            System.out.print(".");
            FileInputStream fichier = new FileInputStream(f);
            wb = new XSSFWorkbook(fichier);
            // iss.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void writeFlux(){
        File f = new File(path+name);
        try {
            System.out.print(".");
            FileOutputStream outFile = new FileOutputStream(f);
            wb.write(outFile);
            outFile.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public List <String> getVariables(){
        initReadFlux();
        List <String> listDeVariables = new ArrayList<String>() ;
        XSSFSheet feuille = wb.getSheet(sheetName);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (Row ligne : feuille) {//parcourir les lignes
            System.out.print("-");
            for (Cell cell : ligne) {//parcourir les colonnes
                System.out.print("-");
                //if( cell.getAddress()!= CellAddress.A1);
                //évaluer le type de la variableColumne
                System.out.print("-");
                boolean isString = formulaEvaluator.evaluateInCell(cell).getCellType()==CellType.STRING?true:false;
                boolean isRequireIndex = cell.getColumnIndex()== variableColumn? true:false;
                if(isString && isRequireIndex) listDeVariables.add(cell.getStringCellValue());
            }
        }
        return listDeVariables;
    }    
    public void saveDatatoSheet(String variable, String docStringList){
        initReadFlux();
        XSSFSheet feuille = wb.getSheet(newSheetForResults?result_sheet:sheetName);
        for (Row row : feuille) {
            if(row.getRowNum()>1){
                Cell cell = row.getCell(newSheetForResults?0:variableColumn);
                if (cell.getStringCellValue().equals(variable)) {
                    Cell cell2 = row.createCell(newSheetForResults?3:resultColumn);
                    if(docStringList.length()>2) cell2.setCellValue(docStringList.substring(0,docStringList.length()-1));
                    writeFlux(); 
                    return;
                }
            }
        }
    }
}
