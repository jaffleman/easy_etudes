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
    private String report = "";
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
    private List <String> segmentList = new ArrayList<>();
    private List <Segments> segList = new ArrayList<>();
    List<String[]> rowDataList = new ArrayList<>(); // liste de données d'entré par ligne



    
    /**
     * @param initData
     */
    public Xlsx(OperatingData initData){
        System.out.println("Reading Xlsx file...");
        File f = new File(initData.getPatn()+initData.getExcelFileName());
        try {
            System.out.print(".");
            FileInputStream fichier = new FileInputStream(f);
            this.wb = new XSSFWorkbook(fichier);
            fichier.close();
        } catch (Exception e) {report += "\n"+e.getMessage()+"\n"+f.getName()+"not found!"; }
        this.sheetName = initData.getSheetName();
        this.resultColumn = initData.getResultColumnNuber();
        this.newSheetForResults = initData.newSheetForResult;
        this.pathName = initData.getPatn()+initData.getExcelFileName();
        XSSFSheet sheet = wb.getSheet(this.sheetName);
        report +="\n Excel file: "+this.pathName;
        report +="\nNew sheet for result: "+(this.newSheetForResults?"yes":"no");

        boolean isFound = false;
        System.out.println("Searching for data...");
        for (Row row : sheet) {
            System.out.print(".");
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
        this.getCodeZoneList();
        this.getVariables();
        if(newSheetForResults){
            System.out.println("Creating new sheet for results...");
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
                for (int i = 2; i <= rowDataList.size()+1; i++ ) {

                    row1 = feuille.createRow(i);
                    cell = row1.createCell(0, CellType.STRING);
                    cell.setCellValue(rowDataList.get(i-2)[1]);
                    cell = row1.createCell(1, CellType.STRING);
                    cell.setCellValue(rowDataList.get(i-2)[2]);
                    cell = row1.createCell(2, CellType.STRING);
                    cell.setCellValue(rowDataList.get(i-2)[0]);
                    cell = row1.createCell(3, CellType.STRING);
                    cell.setCellValue(rowDataList.get(i-2)[3]);
                }
                writeFlux();
            }
        }
        
        
    }



    /**
     * @return
     * return a list of variables from the codeZone column from the Xlsx file
     */
    public void getCodeZoneList() {
        XSSFSheet sheet = wb.getSheet(this.sheetName);
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            if (currentRow.getRowNum()>this.codeZoneRowIndex){
                String[] variableTab = new String[6]; //Stockage des données dans un tableau de 5 elements
                Cell codeZoneCell = currentRow.getCell(codeZoneColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell segmentCell = currentRow.getCell(segmentColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell sousSegmentCell = currentRow.getCell(sousSegmentColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                Cell codeRubriqueHRA = currentRow.getCell(codeRubriqueHRAIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                try {
                    variableTab[0] = segmentCell.getStringCellValue();
                    variableTab[1] = cellValueExtractor(sousSegmentCell);
                    variableTab[2] = cellValueExtractor(codeZoneCell);
                    variableTab[3] = cellValueExtractor(codeRubriqueHRA);
                    variableTab[4] = "";
                    variableTab[5] = Integer.toString(currentRow.getRowNum());
                    rowDataList.add(variableTab);            
                } catch (Exception e) {
                }
            }
        }
        report +="\n number of variable found: "+rowDataList.size();
        report +="\n\n\n";
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
                value = Double.toString(cellule.getNumericCellValue()).split("")[0];
                break;
            default:
                value = "";
                break;
        }
        return value;
    }

    private void getVariables(){
        List <SousSegment> sousSegList = new ArrayList<>();
        List <String[]> zoneList = new ArrayList<>();
        String prevSegName = "null", prevSsegName="null";
        int size = rowDataList.size();
        //String[] maZone=new String[2];
        for (int row = 0; row<size; row++) {
            if (rowDataList.get(row)[0].equals(prevSegName)){
                if (rowDataList.get(row)[1].equals(prevSsegName)){
                    zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(row)[4], rowDataList.get(row)[5]});
                    if (row == size-1){
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList)); 
                        segmentList.add(prevSegName);
                        sousSegList.clear();                      
                    }
                }else{
                    //zoneList.add(maZone);
                    // zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(4)[2], rowDataList.get(row)[5]});
                    sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                    prevSsegName = rowDataList.get(row)[1];
                    zoneList.clear();
                    zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(row)[4], rowDataList.get(row)[5]});
                    //maZone = new String[] {rowDataList.get(i)[0], rowDataList.get(i)[3]};
                    if (row == size-1){
                        // zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(4)[2], rowDataList.get(row)[5]});
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList));                       
                        segmentList.add(prevSegName);
                        sousSegList.clear();                      
                    }
                }
            }else{
                if(row==0) {
                    prevSegName = rowDataList.get(0)[0];
                    prevSsegName = rowDataList.get(0)[1];
                    zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(row)[4], rowDataList.get(row)[5]});
                }else{
                    //zoneList.add(maZone);
                    // zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(4)[2], rowDataList.get(row)[5]});
                    sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                    segmentList.add(prevSegName);
                    segList.add(new Segments(prevSegName, sousSegList));
                    sousSegList.clear();                      
                    prevSsegName = rowDataList.get(row)[1];
                    prevSegName = rowDataList.get(row)[0];
                    zoneList.clear();
                    zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(row)[4], rowDataList.get(row)[5]});
                    //maZone = new String[] {rowDataList.get(i)[0], rowDataList.get(i)[3]};
                    if (row==size-1){
                        // zoneList.add(new String[]{rowDataList.get(row)[2], rowDataList.get(row)[3],rowDataList.get(4)[2], rowDataList.get(row)[5]});
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList));
                        segmentList.add(prevSegName);
                        sousSegList.clear();                      
                    }
                }
            }
        }
        System.out.println("");
    }

    public List<Segments> getSegList() {
        return segList;
    }
    public List<String> getSegmentList() {
        return segmentList;
    }
  
    public void saveDatatoSheet(){
        XSSFSheet feuille = wb.getSheet(this.newSheetForResults?RESULT_SHEET:sheetName);
        for(Segments segment : this.segList){
            for (SousSegment sousSegment : segment.getSsList()) {
                for (String[] variable : sousSegment.getVarList()){
                Row row = feuille.getRow(Integer.parseInt(variable[3]));
                Cell cell = row.createCell(newSheetForResults?4:resultColumn);
                cell.setCellValue(variable[2]);
                }
            }
        }
        writeFlux();
    }
    public String getReport() {
        return report;
    }
}
