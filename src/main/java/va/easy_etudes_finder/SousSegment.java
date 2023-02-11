package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.List;

public class SousSegment {
    private String name = "";
    private List<String[]> xlsxVarList = new ArrayList<>();
    private List<String> docxVarList= new ArrayList<>();
    public SousSegment(String name, List <String[]> variables) {
        this.name = name;
        for (String[] strings : variables) {
            this.xlsxVarList.add(strings);
        }
    }
    public SousSegment(String name2, List<String> stringTabString, int n) {
        this.name = name2;
        for (String strings : stringTabString) {
            this.docxVarList.add(strings);
        }
    }
    public String getName() {
        return name;
    }
    public List<String[]> getXlsxVarList() {
        return xlsxVarList;
    }
    public void add(String data){
        this.xlsxVarList.add(new String[]{data,""});
    }
    public List<String> getDocxVarList() {
        return docxVarList;
    }
}
