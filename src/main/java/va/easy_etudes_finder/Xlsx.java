package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlsx extends Fichier {
    private final String RESULT_SHEET = "result_sheet";
    private String sheetName = "";
    private int resultColumn;
    private XSSFWorkbook wb;
    private boolean newSheetForResults=false;
    private String pathName;
    private int sousSegmentColumnIndex;
    private int segmentColumnIndex;
    private int codeZoneColumnIndex;
    private int codeRubriqueHRAIndex;
    private int codeZoneRowIndex;
    
    public Xlsx(OperatingData initData){
        File f = new File(initData.getPatn()+initData.getExcelFileName());
        try {System.out.print(".");
            FileInputStream fichier = new FileInputStream(f);
            this.wb = new XSSFWorkbook(fichier);
        } catch (Exception e) {e.printStackTrace(); }
        this.sheetName = initData.getSheetName();
        this.resultColumn = initData.getResultColumnNuber();
        this.newSheetForResults = initData.newSheetForResult;
        this.pathName = initData.getPatn()+initData.getExcelFileName();
        XSSFSheet sheet = wb.getSheet(this.sheetName);
        boolean isFound = false;
        for (Row row : sheet) {
            for (Cell cell : row){
                this.codeZoneRowIndex = cell.getRowIndex();
                if (cell.getStringCellValue().equals("Segment")){
                    this.segmentColumnIndex = cell.getColumnIndex();
                }
                if (cell.getStringCellValue().equals("Sous-segment")){
                    this.sousSegmentColumnIndex = cell.getColumnIndex();
                }
                if (cell.getStringCellValue().equals("Code Zone")){
                    this.codeZoneColumnIndex = cell.getColumnIndex();
                }
                if (cell.getStringCellValue().equals("Code rubrique HRA")){
                    this.codeRubriqueHRAIndex = cell.getColumnIndex();
                    isFound = true;
                    break;
                }
            }
            if (isFound) break;
        }
        if(newSheetForResults){
        XSSFSheet feuille = wb.getSheet(RESULT_SHEET);
            if (feuille==null){
                feuille = wb.createSheet(RESULT_SHEET);
                Row row0 = feuille.createRow(0);
                feuille.addMergedRegion(new CellRangeAddress(0, 0, 0, 1 ));
                            
                Cell cell = row0.createCell(0, CellType.STRING);
                cell.setCellValue("SEARCH RESULTS:");
                cell.setCellStyle(CelluleStyle.TitleStyle(wb));
                
                short height = 500;
                row0.setHeight(height);
                
                Row row1 = feuille.createRow(1);
                cell = row1.createCell(0, CellType.STRING);
                cell.setCellValue("Segment:");

                cell = row1.createCell(1, CellType.STRING);
                cell.setCellValue("Sous Segment:");

                cell = row1.createCell(2, CellType.STRING);
                cell.setCellValue("Variables list:");

                cell = row1.createCell(3, CellType.STRING);
                cell.setCellValue("Code rubrique HRA:");

                // cell.setCellStyle(style);
                cell = row1.createCell(4, CellType.STRING);
                cell.setCellValue("found Etudes:");
                // cell.setCellStyle(style);
                List <String[]> myCodeZoneList = getCodeZoneList();
                for (int i = 2; i <= myCodeZoneList.size()+1; i++ ) {

                    row1 = feuille.createRow(i);
                    cell = row1.createCell(0, CellType.STRING);
                    cell.setCellValue(myCodeZoneList.get(i-2)[1]);
                    cell = row1.createCell(1, CellType.STRING);
                    cell.setCellValue(myCodeZoneList.get(i-2)[2]);
                    cell = row1.createCell(2, CellType.STRING);
                    cell.setCellValue(myCodeZoneList.get(i-2)[0]);
                    cell = row1.createCell(3, CellType.STRING);
                    cell.setCellValue(myCodeZoneList.get(i-2)[3]);
                }
                writeFlux();
            }
        }
    }



    /**
     * @return
     * return a list of variables from the codeZone column from the Xlsx file
     */
    public List<String[]> getCodeZoneList() {
        // this.codeZoneRowIndex++;
        XSSFSheet sheet = wb.getSheet(this.sheetName);
        List<String[]> myCodeZoneList = new ArrayList<>();
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            if (!(currentRow.getRowNum()<=this.codeZoneRowIndex)){
                String[] variableTab = new String[4];
                Cell codeZoneCell = currentRow.getCell(codeZoneColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                Cell segmentCell = currentRow.getCell(segmentColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                Cell sousSegmentCell = currentRow.getCell(sousSegmentColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                Cell codeRubriqueHRA = currentRow.getCell(codeRubriqueHRAIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                try {
                    variableTab[0] = codeZoneCell.getStringCellValue();
                    variableTab[1] = segmentCell.getStringCellValue();
                    variableTab[2] = sousSegmentCell.getStringCellValue();
                    variableTab[3] = codeRubriqueHRA.getStringCellValue();
                    myCodeZoneList.add(variableTab);            
                } catch (Exception e) {
                }
            }
        }
        return myCodeZoneList;
    }





    /**
     * write information into the Xlsx file
     */
    private void writeFlux(){
        File f = new File(this.pathName);
        try {
            System.out.print(".");
            FileOutputStream outFile = new FileOutputStream(f);
            wb.write(outFile);
            outFile.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }





  
    public void saveDatatoSheet(List <String[]> datatList){
        XSSFSheet feuille = wb.getSheet(this.newSheetForResults?RESULT_SHEET:sheetName);
        for(String[] data : datatList){
        for (Row row : feuille) {
            if((newSheetForResults && row.getRowNum()>1) || !newSheetForResults){
                Cell cell = row.getCell(this.newSheetForResults?2:codeZoneColumnIndex);
                if (cell != null){
                    if (cell.getStringCellValue().equals(data[0])) {
                        Cell cell2 = row.createCell(newSheetForResults?4:resultColumn);
                        String dataToCell = data[1].length()>2?data[1].substring(0,data[1].length()-1):"";
                        if(data[1].length()>2) cell2.setCellValue(dataToCell);
                        break;
                    }
                }
            }
        }}
        writeFlux(); 
    }
}
