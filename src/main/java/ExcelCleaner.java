import java.io.IOException;

/**
 * Created by Prananda Ramadhan on 22/01/2016.
 */
public class ExcelCleaner {
    public static void main(String[] args) throws IOException {
        POI poi = new POI("Contacts.xlsx");
        poi.readExcel();
        //poi.createExcelContact("Data Contacts.xlsx", "Contact");
        //poi.createExcelDataContact("Data Contacts Main.xlsx", "Contact");
    }
}
