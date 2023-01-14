package va.easy_etudes_finder;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CelluleStyle{
    public static XSSFCellStyle TitleStyle(XSSFWorkbook workbook){
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setItalic(true);
    
        // Font Height
        font.setFontHeightInPoints((short) 16);
    
        // Font Color
        font.setColor(Font.COLOR_NORMAL);
    
        // Style
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
    
        return style;
    }
}
