import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Main {

    public static final String rulesPath = "./exampleRules.xlsx";
    public static final String testingPath = "./exampleReport.xlsx";

    public static Workbook testingBook;
    public static Sheet testingSheet;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        System.out.println("Hello World");

        Workbook workbook = WorkbookFactory.create(new File(rulesPath));
        testingBook = WorkbookFactory.create(new File(testingPath));
        testingSheet = testingBook.getSheetAt(0);

        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();



        Parser p = new Parser();




        for(Row row: sheet) {
            Parser.Node n = null;
            for(Cell cell: row) {
                if(cell.getCellTypeEnum() == CellType.FORMULA) {
                    n = p.parseCell(cell.getCellFormula());
                }
                if(n != null) {
                    System.out.println(n.execute() + " == " + testingSheet.getRow(row.getRowNum()).getCell(cell.getColumnIndex()));
                }
            }
            System.out.println();
        }




    }






}
