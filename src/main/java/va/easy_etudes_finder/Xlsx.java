package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.Arrays;
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
    private int fsfIndex;
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
        File f = new File("/home/Jaffleman/Documents/banque-docx/xlsx/"+initData.getExcelFileName());
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
                }
                if (cell.getStringCellValue().equals("FSF actuel")){
                    this.fsfIndex= cell.getColumnIndex();
                    isFound = true;
                    break;
                }
            }
            if (isFound) break;
        }
        this.getCodeZoneList();
        this.getVariables();        
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
                String[] variableTab = new String[9]; //Stockage des données dans un tableau de 5 elements
                Cell segmentCell = currentRow.getCell(segmentColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell sousSegmentCell = currentRow.getCell(sousSegmentColumnIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                Cell codeZoneCell = currentRow.getCell(codeZoneColumnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell codeRubriqueHRA = currentRow.getCell(codeRubriqueHRAIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                Cell fsfActuel = currentRow.getCell(fsfIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                try {
                    variableTab[0] = segmentCell.getStringCellValue();
                    variableTab[1] = cellValueExtractor(sousSegmentCell);
                    variableTab[2] = cellValueExtractor(codeZoneCell);
                    variableTab[3] = cellValueExtractor(codeRubriqueHRA);
                    variableTab[4] = cellValueExtractor(fsfActuel);
                    variableTab[5] = Integer.toString(currentRow.getRowNum());
                    rowDataList.add(variableTab);            
                } catch (Exception e) {
                    //report +="\nSomething goes wong in 'getCodeZoneList()' (SheetRow: "+currentRow.getRowNum()+")";
                }
            }
        }
        report +="\n number of variable found: "+rowDataList.size();
        report +="\n";
    }





    /**
     * write information into the Xlsx file
     */
    private void writeFlux(){
        System.out.print("\nFinalising...");
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
                Double dNumber = cellule.getNumericCellValue();
                int iNumber = dNumber.intValue();
                value = Integer.toString(iNumber);
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
            String segment = rowDataList.get(row)[0];
            String sousSegment = rowDataList.get(row)[1];
            String codeZone = rowDataList.get(row)[2];
            String codeBHA =rowDataList.get(row)[3];
            String fsfActuel = rowDataList.get(row)[4];
            String rowIndex = rowDataList.get(row)[5];
            if (segment.equals(prevSegName)){
                if (sousSegment.equals(prevSsegName)){
                    zoneList.add(new String[]{codeZone,codeBHA, fsfActuel,"", "", "", rowIndex});
                    if (row == size-1){
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList)); 
                        segmentList.add(prevSegName);
                        sousSegList.clear();                      
                    }
                }else{
                    //zoneList.add(maZone);
                    // zoneList.add(new String[]{codeZone, fsfActuel,rowDataList.get(4)[2], overFoundFsf});
                    sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                    prevSsegName = sousSegment;
                    zoneList.clear();
                    zoneList.add(new String[]{codeZone, codeBHA,fsfActuel,"", "", "",rowIndex});
                    //maZone = new String[] {rowDataList.get(i)[0], rowDataList.get(i)[3]};
                    if (row == size-1){
                        // zoneList.add(new String[]{codeZone, fsfActuel,rowDataList.get(4)[2], overFoundFsf});
                        sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                        zoneList.clear();
                        segList.add(new Segments(prevSegName, sousSegList));                       
                        segmentList.add(prevSegName);
                        sousSegList.clear();                      
                    }
                }
            }else{
                if(row==0) {
                    prevSegName = segment;
                    prevSsegName = sousSegment;
                    zoneList.add(new String[]{codeZone,codeBHA, fsfActuel,"", "", "",rowIndex});
                }else{
                    //zoneList.add(maZone);
                    // zoneList.add(new String[]{codeZone, fsfActuel,rowDataList.get(4)[2], overFoundFsf});
                    sousSegList.add(new SousSegment(prevSsegName, new ArrayList<>(zoneList)));
                    segmentList.add(prevSegName);
                    segList.add(new Segments(prevSegName, sousSegList));
                    sousSegList.clear();                      
                    prevSsegName = sousSegment;
                    prevSegName = segment;
                    zoneList.clear();
                    zoneList.add(new String[]{codeZone,codeBHA, fsfActuel,"", "", "",rowIndex});
                    //maZone = new String[] {rowDataList.get(i)[0], rowDataList.get(i)[3]};
                    if (row==size-1){
                        // zoneList.add(new String[]{codeZone, fsfActuel,rowDataList.get(4)[2], overFoundFsf});
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
    private void compatator(){
        for(Segments segment : this.segList){
            for (SousSegment sousSegment : segment.getSsList()) {
                for (String[] variable : sousSegment.getXlsxVarList()){
                    String unfoundValue ="";
                    String unfoundValue2 ="";
                    for (String easyEtudesFoundedDCD : variable[3].split(";")){
                        if (!variable[2].contains(easyEtudesFoundedDCD)) unfoundValue += easyEtudesFoundedDCD+";";
                    }
                    for (String fsfActuelElement : variable[2].split(";")){
                        if (!variable[3].contains(fsfActuelElement)) unfoundValue2 += fsfActuelElement+";";
                    }
                    variable[5] = unfoundValue;
                    variable[4] = unfoundValue2;
                }
            }
        }
    }
    public void saveDatatoSheet(){
        this.compatator();
        XSSFSheet feuille = wb.getSheet(this.newSheetForResults?RESULT_SHEET:sheetName);


        if(newSheetForResults){
            System.out.println("Creating new sheet for results...");
            if (feuille==null){
                feuille = wb.createSheet(RESULT_SHEET);
                Row row = feuille.createRow(0);
                feuille.addMergedRegion(new CellRangeAddress(0, 0, 0, 1 ));
                            
                Cell cell = row.createCell(0, CellType.STRING);
                cell.setCellValue("SEARCH RESULTS:");
                cell.setCellStyle(CelluleStyle.TitleStyle(wb));
                
                // short height = 500;
                // row.setHeight(height);
                
                row = feuille.createRow(1);
                cell = row.createCell(0, CellType.STRING);
                cell.setCellValue("Segment:");

                cell = row.createCell(1, CellType.STRING);
                cell.setCellValue("Sous Segment:");

                cell = row.createCell(2, CellType.STRING);
                cell.setCellValue("Code Zone:");

                cell = row.createCell(3, CellType.STRING);
                cell.setCellValue("Code rubrique HRA:");

                cell = row.createCell(4, CellType.STRING);
                cell.setCellValue("FSF actuel:");
                
                cell = row.createCell(5, CellType.STRING);
                cell.setCellValue("EasyEtudes result:");

                cell = row.createCell(6, CellType.STRING);
                cell.setCellValue("missing etudes:");

                cell = row.createCell(7, CellType.STRING);
                cell.setCellValue("New found etudes:");

            }
        }
        for(Segments segment : this.segList){
        System.out.print(".");
            for (SousSegment sousSegment : segment.getSsList()) {
                for (String[] variable : sousSegment.getXlsxVarList()){
                    Row row = newSheetForResults?feuille.createRow((Integer.parseInt(variable[6]))-2):
                        feuille.getRow(Integer.parseInt(variable[6]));
                    if (newSheetForResults){
                        Cell cell = row.createCell(0, CellType.STRING);
                        cell.setCellValue(segment.getName());
                        cell = row.createCell(1, CellType.STRING);
                        cell.setCellValue(sousSegment.getName());
                        cell = row.createCell(2, CellType.STRING);
                        cell.setCellValue(variable[0]);
                        cell = row.createCell(3, CellType.STRING);
                        cell.setCellValue(variable[1]);
                        cell = row.createCell(4, CellType.STRING);
                        cell.setCellValue(variable[2]);
                    }
                    Cell cell = row.createCell(newSheetForResults?5:resultColumn);
                    Cell delta = row.createCell(newSheetForResults?6:resultColumn+1);
                    Cell undelta = row.createCell(newSheetForResults?7:resultColumn+2);
                    cell.setCellValue(variable[3]);
                    delta.setCellValue(variable[4]);
                    undelta.setCellValue(variable[5]);
                }
            }
        }
        writeFlux();
    }
    public void workOnDcd(List<Docx2> docx2List){
        for (Docx2 docx : docx2List){
            System.out.print(".");
        for (Segments docxSegment : docx.getSegList()){
            int index = segmentList.indexOf(docxSegment.getName());
            if (index == -1) {
                report+= "\nUnknow Segment: "+docxSegment.getName()+" in docx: "+docx.getshortName();
            }else{
            Segments excelSegment = segList.get(index);
            for (SousSegment docxSousSegment : docxSegment.getSsList()) {
                Boolean sousSegExist = false;
                for (SousSegment excelSousSegment : excelSegment.getSsList()) {
                    String excelSousSegName = excelSousSegment.getName();
                    String docxSousSegName = docxSousSegment.getName();
                    if (docxSousSegName.equals(excelSousSegName)) {
                        sousSegExist = true;
                        for (String[] var : excelSousSegment.getXlsxVarList()) {
                            for (String docxVar : docxSousSegment.getDocxVarList()) {
                                if( docxVar.equals(var[0])){
                                    if(!(Arrays.asList(var[3].split(";")).contains(docx.getshortName()))) 
                                        var[3]+=docx.getshortName()+";";
                                }
                            }
                            
                        }
                    }
                }
                if (!sousSegExist) {
                    for (SousSegment excelSousSegment : excelSegment.getSsList()) {
                        for (String[] var : excelSousSegment.getXlsxVarList()) {
                            for (String docxVar : docxSousSegment.getDocxVarList()) {
                                if( docxVar.equals(var[0])){
                                    if(!(Arrays.asList(var[3].split(";")).contains(docx.getshortName()))) 
                                        var[3]+=docx.getshortName()+";";
                                }
                            }
                        }   
                    } 
                     
                }

            }
            }
        }}
    }
    public String getReport() {
        return report;
    }
    public void creatReportSheet(List <Docx2> docx2List){
        System.out.println(".");
        XSSFSheet feuille = wb.createSheet("report");
        Row row = feuille.createRow(0);
        feuille.addMergedRegion(new CellRangeAddress(0, 0, 0, 1 ));
                    
        Cell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("EasyEtudes report:");
        cell.setCellStyle(CelluleStyle.TitleStyle(wb));
        
        row = feuille.createRow(1);
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Docx filename:");

        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Stat:");

        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Report:");

        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("Extract text:");

        cell = row.createCell(4, CellType.STRING);
        cell.setCellValue("Segment:");
        
        cell = row.createCell(5, CellType.STRING);
        cell.setCellValue("Sous-segment:");

        cell = row.createCell(6, CellType.STRING);
        cell.setCellValue("Code-Zone:");
        int low = 2;
        int hight = 2;
        int ssLow =2;
        int sLow = 2;
        for (Docx2  docx2 : docx2List) {
            System.out.print(".");
            if(docx2.getshortName().equals("GRADE")){
                System.out.println("GRADE");
            }
            if (docx2.getSegList().size()>0) {
                for(Segments segment :docx2.getSegList()){
                    for (SousSegment ssegment : segment.getSsList()){
                        for (String varString: ssegment.getDocxVarList()){
                            row = feuille.createRow(hight);
                            cell = row.createCell(6, CellType.STRING);
                            cell.setCellValue(varString);
                            hight++;
                        }
                        if (low<hight-1){feuille.addMergedRegion(new CellRangeAddress(low, hight-1 , 5, 5 ));}
                        row=feuille.getRow(low);
                        cell = row.createCell(5, CellType.STRING);
                        cell.setCellValue(ssegment.getName());
                        low=hight;
                    }
                    if (ssLow<hight-1){feuille.addMergedRegion(new CellRangeAddress(ssLow, hight-1 , 4, 4 ));}
                    row=feuille.getRow(ssLow);
                    cell = row.createCell(4, CellType.STRING);
                    cell.setCellValue(segment.getName());
                    ssLow=hight;
                }
            }else{
                hight++;
            }
            if (sLow<hight-1){
                feuille.addMergedRegion(new CellRangeAddress(sLow, hight-1 , 0, 0 )); 
                feuille.addMergedRegion(new CellRangeAddress(sLow, hight-1 , 1, 1 )); 
                feuille.addMergedRegion(new CellRangeAddress(sLow, hight-1 , 2, 2 )); 
                feuille.addMergedRegion(new CellRangeAddress(sLow, hight-1 , 3, 3 )); 
                row=feuille.getRow(sLow);
            }else{
                row = feuille.createRow(low);
            }
                cell = row.createCell(0, CellType.STRING);
                cell.setCellValue(docx2.getshortName());
                cell = row.createCell(1, CellType.STRING);
                cell.setCellValue(docx2.getStatus());
                cell = row.createCell(2, CellType.STRING);
                cell.setCellValue(docx2.getReport());
                cell = row.createCell(3, CellType.STRING);
                cell.setCellValue(docx2.getExtractText());
            sLow = hight; 
            
        }}}
