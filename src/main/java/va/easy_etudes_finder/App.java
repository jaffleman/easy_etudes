package va.easy_etudes_finder;

import java.io.*;
import org.apache.poi.xwpf.extractor.*;
import org.apache.poi.extractor.*;

public class App {
    public static void main(String[] args) throws IOException {

        File f=  new File("/home/Jaffleman/Documents/COVID-19 - Attestation-sur-l-honneur.docx");
        FileInputStream iss = null;
                     try{
        POITextExtractor textExtractor = ExtractorFactory.createExtractor(f);
         XWPFWordExtractor wordExtractor = (XWPFWordExtractor) textExtractor;
                    String contenu = wordExtractor.getText();
                    System.out.println(contenu);
        }
        catch (Exception e) {
                          e.printStackTrace();
                        }
        finally {
            if (iss != null) iss.close();
        }

    }

}
