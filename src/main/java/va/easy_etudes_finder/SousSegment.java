package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.List;

public class SousSegment {
    private String name = "";
    private List<String[]> varList = new ArrayList<>();
    public SousSegment(String name, List <String[]> variables) {
        this.name = name;
        for (String[] strings : variables) {
            this.varList.add(strings);
        }
    }
    public String getName() {
        return name;
    }
    public List<String[]> getVarList() {
        return varList;
    }
}
