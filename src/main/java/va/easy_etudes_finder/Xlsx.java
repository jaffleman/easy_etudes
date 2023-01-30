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
// import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlsx {
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
    
    /**
     * @param initData
     */
    public Xlsx(OperatingData initData){
        File f = new File(initData.getPatn()+initData.getExcelFileName());
        try {System.out.print(".");
            FileInputStream fichier = new FileInputStream(f);
            this.wb = new XSSFWorkbook(fichier);
            fichier.close();
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
                    
                    variableTab[0] = codeZoneCell.getStringCellValue().toString();
                    variableTab[1] = segmentCell.getStringCellValue().toString();
                    variableTab[2] = sousSegmentCell.getStringCellValue().toString();
                    variableTab[3] = codeRubriqueHRA.getStringCellValue().toString();
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


    private String cellValueExtractor(Cell cellule){
        String value;
        if (cellule==null) return "";
        switch (cellule.getCellType()) {
            case STRING:
                value = cellule.getStringCellValue();
                break;
            case NUMERIC:
                value = Double.toString(cellule.getNumericCellValue());               
                break;
            default:
                value = "";
                break;
        }
        return value;
    }

    public List <Segments> getVariables(){
        XSSFSheet sheet = wb.getSheet(this.sheetName);
        List<String[]> myCodeZoneList = new ArrayList<>();
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            if (currentRow.getRowNum()>this.codeZoneRowIndex){
                String[] variableTab = new String[5];
                Cell codeZoneCell = currentRow.getCell(codeZoneColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell segmentCell = currentRow.getCell(segmentColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell sousSegmentCell = currentRow.getCell(sousSegmentColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                Cell codeRubriqueHRA = currentRow.getCell(codeRubriqueHRAIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                try {
                    variableTab[1] = segmentCell.getStringCellValue();
                    variableTab[0] = cellValueExtractor(codeZoneCell);
                    variableTab[2] = cellValueExtractor(sousSegmentCell);
                    variableTab[3] = cellValueExtractor(codeRubriqueHRA);
                    variableTab[4] = Integer.toString(currentRow.getRowNum()) ;
                    myCodeZoneList.add(variableTab);            
                } catch (Exception e) {
                }
            }
        }
        List <Segments> segList = new ArrayList<>();
        List <SousSegment> sousSegList = new ArrayList<>();
        List <String[]> zoneList = new ArrayList<>();
        String prevSegName = "null", prevSsegName="null";
        int size = myCodeZoneList.size();
        //String[] maZone=new String[2];
        for (int i = 0; i<size; i++) {
            if (myCodeZoneList.get(i)[1].equals(prevSegName)){
                if (myCodeZoneList.get(i)[2].equals(prevSsegName)){
                    zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                    //maZone = new String[] {myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3]};
                    if (i == size-1){
                        zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList)); 
                        sousSegList.clear();                      
                    }
                }else{
                    //zoneList.add(maZone);
                    zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                    sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                    prevSsegName = myCodeZoneList.get(i)[2];
                    zoneList.clear();
                    zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                    //maZone = new String[] {myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3]};
                    if (i == size-1){
                        zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList));                       
                        sousSegList.clear();                      
                    }
                }
            }else{
                if(i==0) {
                    prevSegName = myCodeZoneList.get(0)[1];
                    prevSsegName = myCodeZoneList.get(0)[2];
                    zoneList.add(new String[]{myCodeZoneList.get(0)[0], myCodeZoneList.get(0)[3], myCodeZoneList.get(i)[4]});
                }else{
                    //zoneList.add(maZone);
                    zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                    sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                    segList.add(new Segments(prevSegName, sousSegList));
                    sousSegList.clear();                      
                    prevSsegName = myCodeZoneList.get(i)[2];
                    prevSegName = myCodeZoneList.get(i)[1];
                    zoneList.clear();
                    zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                    //maZone = new String[] {myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3]};
                    if (i==size-1){
                        zoneList.add(new String[]{myCodeZoneList.get(i)[0], myCodeZoneList.get(i)[3], myCodeZoneList.get(i)[4]});
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList));
                        sousSegList.clear();                      
                    }
                }
            }
        }
        return segList;
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
