import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Main {

    public static final String rulesPath = "./adamsTestWorkbook.xlsx";
    public static final String testingPath = "./exampleReport.xlsx";

    public static Workbook testingBook;
    public static Sheet testingSheet;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        System.out.println("Hello World");

        Workbook workbook = WorkbookFactory.create(new File(rulesPath));
        testingBook = WorkbookFactory.create(new File(testingPath));
        testingSheet = testingBook.getSheetAt(0);

        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();



        Parser p = new Parser();
        Parser.Node n = p.parseCell("5+7*9");

        System.out.println(n.execute());


        for(Row row: sheet) {

            for(Cell cell: row) {
                System.out.println(cell);
                System.out.println(cell.getCellTypeEnum());
            }
            System.out.println();
        }




    }






}
