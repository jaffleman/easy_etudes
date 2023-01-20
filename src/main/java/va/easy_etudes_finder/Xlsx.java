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
    private int variableSeize;
    private List<String> varables;

    public Xlsx(OperatingData initData){
        super.name = initData.getExcelFileName();
        super.path = initData.getPatn();  
        this.sheetName = initData.getSheetName();
        this.variableColumn = initData.getVariableColumnNumber(); 
        this.resultColumn = initData.getResultColumnNuber();
        varables = this.initVariables();
        if (this.resultColumn==-1){
            this.newSheetForResults=true;
            XSSFSheet feuille = wb.getSheet(result_sheet);
            if (feuille==null){
                feuille = wb.createSheet(result_sheet);
                Row row0 = feuille.createRow(0);
                feuille.addMergedRegion(new CellRangeAddress(0, 0, 0, 1 ));
                            
                Cell cell = row0.createCell(0, CellType.STRING);
                cell.setCellValue("SEARCH RESULTS:");
                cell.setCellStyle(CelluleStyle.TitleStyle(wb));
                
                short height = 500;
                row0.setHeight(height);
                
                Row row1 = feuille.createRow(1);
                cell = row1.createCell(0, CellType.STRING);
                cell.setCellValue("Variables list:");
                // cell.setCellStyle(style);
                cell = row1.createCell(1, CellType.STRING);
                cell.setCellValue("found Etudes:");
                // cell.setCellStyle(style);
                for (int i = 2; i <= varables.size()+1; i++ ) {
                    row1 = feuille.createRow(i);
                    cell = row1.createCell(0, CellType.STRING);
                    cell.setCellValue(varables.get(i-2));
                }
                writeFlux();
            } else {
            }
        }
    }

    public Xlsx(String filePath, String fileName, String sheetName, int variableColumn, int rColumn) {
        // same sheet
        super();
        super.name = fileName;
        super.path = filePath;  
        this.sheetName = sheetName;
        this.variableColumn = variableColumn; 
        varables = this.initVariables();
        
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
    
    public int getSeize(){
        return this.variableSeize;
    }
    public List <String> getVariables(){
        return varables;
    }
    public List <String> initVariables(){
        System.out.print(".");
        initReadFlux();
        List <String> listDeVariables = new ArrayList<String>() ;
        XSSFSheet feuille = wb.getSheet(this.sheetName);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (Row ligne : feuille) {//parcourir les lignes
            System.out.print(".");
            //boolean isVisible = !ligne.getZeroHeight();
            Cell cell = ligne.getCell(variableColumn);
            
            if (cell!=null&&//isVisible&&
                formulaEvaluator.evaluateInCell(cell).getCellType()==CellType.STRING
            ) {
                String cellValue = cell.getStringCellValue();
                if( 
                    !cell.getStringCellValue().equals("Debut")&&
                    !cell.getStringCellValue().equals("DEBUT")&&
                    !cell.getStringCellValue().equals("debut")&&
                    !cell.getStringCellValue().equals("FIN")&&
                    !cell.getStringCellValue().equals("Fin")&&
                    !cell.getStringCellValue().equals("fin")&&
                    !cell.getStringCellValue().equals("SUBTY")
                ) listDeVariables.add(cellValue);
            }
        }
        this.variableSeize = listDeVariables.size();
        formulaEvaluator.clearAllCachedResultValues();
        return listDeVariables;
    }    
    public void saveDatatoSheet(List <String[]> datatList){
        initReadFlux();
        XSSFSheet feuille = wb.getSheet(newSheetForResults?result_sheet:sheetName);
        for(String[] data : datatList){
        for (Row row : feuille) {
            if((newSheetForResults && row.getRowNum()>1) || !newSheetForResults){
                Cell cell = row.getCell(newSheetForResults?0:variableColumn);
                if (cell != null){
                    if (cell.getStringCellValue().equals(data[0])) {
                        Cell cell2 = row.createCell(newSheetForResults?1:resultColumn);
                        if(data[1].length()>2) cell2.setCellValue(data[1].substring(0,data[1].length()-1));
                        System.out.print(".");
                        break;
                    }
                }
            }
        }}
        writeFlux(); 
    }
}
