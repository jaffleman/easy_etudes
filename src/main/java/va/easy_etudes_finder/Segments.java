package va.easy_etudes_finder;

import java.util.ArrayList;
import java.util.List;

public class Segments {
    private String name;
    private List<SousSegment> ssList = new ArrayList<>();
    public Segments(String name, List <SousSegment> ssegments) {
        this.name = name;
        List <SousSegment> ssegmentList = ssegments;
        for (SousSegment sousSegment : ssegmentList) {
            this.ssList.add(sousSegment);
        }
    }
    public String getName() {
        return name;
    }
    public List<SousSegment> getSsList() {
        return ssList;
    }
}
