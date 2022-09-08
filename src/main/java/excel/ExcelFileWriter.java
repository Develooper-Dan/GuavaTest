package excel;

import com.google.common.collect.Table;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;

public class ExcelFileWriter {

    static void createXLSXFromTable(Table<Integer, Integer, String> table) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        for (Table.Cell<Integer, Integer, String> cell : table.cellSet()) {
            sheet.createRow(cell.getRowKey()).createCell(cell.getColumnKey()).setCellValue(cell.getValue());
        }
        try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    static void createCSVFromTable(Table<Integer, Integer, String> table) {
        try (CSVPrinter printer = new CSVPrinter(new FileWriter("excel.csv"), CSVFormat.EXCEL)) {
            printer.printRecords(table.values());
            /*for(Table.Cell<Integer, Integer, String> cell : table.cellSet()){
               printer.print(cell.getValue());
            }*/
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Excel3000 excel = new Excel3000();
        excel.createTable();
        excel.setCell("a2", "0.5");
        excel.setCell("c3", "17");
        excel.setCell("BA12", "=$a2^2 - $c3");
        Table<Integer, Integer, String> calculatedTable = excel.evaluate();
        createXLSXFromTable(calculatedTable);
        createCSVFromTable(calculatedTable);
    }
}
