/**
 * Created by Prananda Ramadhan on 22/01/2016.
 */
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class POI {
    private Map<Integer, Map<String, Cell>> data;
    private Map<Integer, List<Cell>> contactData;
    private int totalRows;
    private int totalCells;
    private FileInputStream excel;
    private Sheet sheet;
    private Workbook workbook;

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public void setWorkbook(String path) throws IOException {
        if (path.endsWith("xlsx")) {
            this.workbook = new XSSFWorkbook(getExcel());
        } else if (path.endsWith("xls")) {
            this.workbook = new HSSFWorkbook(getExcel());
        }
        else {
            throw new IllegalArgumentException("It is not an excel file");
        }
    }

    public FileInputStream getExcel() {
        return excel;
    }

    public void setExcel(FileInputStream excel, String uri) throws IOException {
        this.excel = new FileInputStream(new File(uri));
    }

    public int getTotalRows() {
        return totalRows;
    }

    public void setTotalRows(int totalRows) {
        this.totalRows = totalRows;
    }

    public int getTotalCells() {
        return totalCells;
    }

    public void setTotalCells(int totalCells) {
        this.totalCells = totalCells;
    }

    public Map<Integer, Map<String, Cell>> getData() {
        return data;
    }

    public void setData(Map<Integer, Map<String, Cell>> data) {
        this.data = data;
    }

    public Map<Integer, List<Cell>> getContactData() {
        return contactData;
    }

    public void setContactData(Map<Integer, List<Cell>> contactData) {
        this.contactData = contactData;
    }

    public POI(String uri) throws IOException {
        setExcel(this.excel, uri);
        setWorkbook(uri);
        setSheet(getWorkbook().getSheetAt(0));
        setTotalRows(getSheet().getPhysicalNumberOfRows());
        setTotalCells(getSheet().getRow(0).getPhysicalNumberOfCells());
    }

    public void readExcel() throws IOException {
        System.out.println("total rows: " + getTotalRows() + ", total cells: " + getTotalCells());
        Map<Integer, List<Cell>> data = new HashMap<>();
        Map<Integer, Map<String, Cell>> contact = new HashMap<>();
        //
        Map<String, Cell> contactValue = new HashMap<>();
        List<Cell> dataValue = new ArrayList<>();
        //
        Cell cell;
        int contactId;

        for (int rowNumber = 1; rowNumber < getTotalRows() ; rowNumber++) {
            Row row = getSheet().getRow(rowNumber);
            contactId = (int) row.getCell(2).getNumericCellValue();
            if (row == null) {
                continue;
            } else {
                for (int cellNumber = 0; cellNumber < getTotalCells() ; cellNumber++) {
                    cell = row.getCell(cellNumber);
                    // 0 1 11 13 15
                    switch (cellNumber) {
                        case 0:
                            contactValue.put("Business Fax", cell);
                            break;
                        case 1:
                            contactValue.put("Business Phone", cell);
                            break;
                        case 11:
                            contactValue.put("Home Phone", cell);
                            break;
                        case 13:
                            contactValue.put("Mobile Phone", cell);
                            break;
                        case 15:
                            contactValue.put("Pager", cell);
                            break;
                    }
                    if (cellNumber == 0 || cellNumber == 1 || cellNumber == 11 || cellNumber == 13 || cellNumber == 15) {
                        contact.put(contactId, contactValue);
                    }
                    else if(cellNumber == 2 || cellNumber == 6 || cellNumber == 10) {
                        continue;
                    }
                    else {
                        dataValue.add(cell);
                        data.put(contactId, dataValue);
                    }
                }
            }
            dataValue = new ArrayList<>();
            contactValue = new HashMap<>();
        }
        setData(contact);
        setContactData(data);
        getWorkbook().close();
        getExcel().close();
    }

    public void createExcelContact(String fileName, String sheetName) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(sheetName);
        int rowIndex = 0;
        for (Map.Entry<Integer, Map<String, Cell>> eachRow : getData().entrySet()) {
            if (eachRow.getValue() instanceof Map) {
                Iterator iterator = eachRow.getValue().entrySet().iterator();
                while (iterator.hasNext()) {
                    Map.Entry eachCell = (Map.Entry) iterator.next();
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(1).setCellValue(eachRow.getKey());
                    row.createCell(2).setCellValue(eachCell.getKey().toString());
                    if (eachCell.getValue() == null) {
                        row.createCell(3).setCellValue("");
                    } else {
                        row.createCell(3).setCellValue(eachCell.getValue().toString());
                    }
                }
            }
        }
        FileOutputStream excelOut = new FileOutputStream(fileName);
        wb.write(excelOut);
        excelOut.close();
    }

    public void createExcelDataContact(String fileName, String sheetName) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(sheetName);

        int index = 0;
        for (Map.Entry<Integer, List<Cell>> eachRow : getContactData().entrySet()) {
            Row row = sheet.createRow(index);
            row.createCell(0).setCellValue(eachRow.getKey());
            for (int i = 1; i < eachRow.getValue().size(); i++) {
                if (eachRow.getValue().get(i) == null) {
                    row.createCell(i).setCellValue("");
                } else {
                    row.createCell(i).setCellValue(eachRow.getValue().get(i).toString());
                }
            }
            index++;
        }

        FileOutputStream excelOut = new FileOutputStream(fileName);
        wb.write(excelOut);
        excelOut.close();
    }

    public Map<Object, Object> mapValue() {
        return mapValue();
    }

}
