package va.easy_etudes_finder;

import java.util.Arrays;
import java.util.List;

public class Converter {
    public static int convertion(String lettre){
        if(lettre.length()>0){
            int base = 26;
            int result = 0;
            String[] tabChar = new String[]{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
            List<String> charList = Arrays.asList(tabChar);
            String[] splitLettre = lettre.split("");
            for(int i=0; i< splitLettre.length/2; i++){
                String tmp = splitLettre[i];
                splitLettre[i] = splitLettre[splitLettre.length-i-1];
                splitLettre[splitLettre.length-i-1] = tmp;
            }
            for (int i=0; i<splitLettre.length; i++) {
                result += (Math.pow(base,i))*(charList.indexOf(splitLettre[i])+1);
            }
            return result-1;
        }else return -1;
    }
}
