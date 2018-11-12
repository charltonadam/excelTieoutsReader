import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
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

        Workbook rulesBook = WorkbookFactory.create(new File(rulesPath));
        testingBook = WorkbookFactory.create(new File(testingPath));
        testingSheet = testingBook.getSheetAt(0);

        Sheet rulesSheet = rulesBook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();

        Workbook resultBook = new XSSFWorkbook();
        CreationHelper createHelper = resultBook.getCreationHelper();

        Sheet resultSheet = resultBook.createSheet();

        Font errorFont = resultBook.createFont();
        errorFont.setBold(true);
        errorFont.setColor(IndexedColors.RED.getIndex());

        CellStyle errorCellStyle = resultBook.createCellStyle();
        errorCellStyle.setFont(errorFont);



        Parser p = new Parser();




        for(Row row: rulesSheet) {
            Row errorRow = resultSheet.createRow(row.getRowNum());
            Parser.Node n = null;
            for(Cell cell: row) {
                if(cell.getCellTypeEnum() == CellType.FORMULA) {
                    n = p.parseCell(cell.getCellFormula());
                }
                if(n != null) {
                    double result = testingSheet.getRow(row.getRowNum()).getCell(cell.getColumnIndex()).getNumericCellValue() - n.execute();
                    System.out.println(n.execute() + " == " + testingSheet.getRow(row.getRowNum()).getCell(cell.getColumnIndex()));
                    Cell errorCell = errorRow.createCell(cell.getColumnIndex());
                    if(result != 0) {
                        //Houston, we have a problem, mark it on the results sheet

                        errorCell.setCellValue(result);
                        errorCell.setCellStyle(errorCellStyle);
                    }
                }
            }
        }

        rulesBook.close();
        testingBook.close();


        FileOutputStream fileOut = new FileOutputStream("results.xlsx");
        resultBook.write(fileOut);
        fileOut.close();

        resultBook.close();




    }






}
