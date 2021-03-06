package excelToCsv;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Utils {

    static protected int getLastColNum(Sheet sheet) {
        
        int lastColNo = 0;
        
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if(row == null) {
                continue;
            }
            int no = row.getLastCellNum();
            if(lastColNo < no) {
                lastColNo = no;
            }
        }
        return lastColNo;
    }    
}
