
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import static org.apache.poi.ss.usermodel.CellType.*;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class CopyExcel {

    public static void main(String[] args) {

        // Read xlsx file
        XSSFWorkbook oldWorkbook = null;
        try {
            oldWorkbook = (XSSFWorkbook) WorkbookFactory.create(new File("dovamo.xlsx"));
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        final XSSFWorkbook newWorkbook = new XSSFWorkbook();

        // Copy style source
        final StylesTable oldStylesSource = oldWorkbook.getStylesSource();
        final StylesTable newStylesSource = newWorkbook.getStylesSource();
        long startTime = System.currentTimeMillis();
//        oldStylesSource.getFonts().forEach(font -> newStylesSource.putFont(font, true));
//        oldStylesSource.getFills().forEach(fill -> newStylesSource.putFill(new XSSFCellFill()));
//        oldStylesSource.getBorders()
//                .forEach(border -> newStylesSource.putBorder(new XSSFCellBorder(border.getCTBorder())));

        // Copy sheets
       // for (int sheetNumber = 0; sheetNumber < oldWorkbook.getShgetNumberOfSheets(); sheetNumber++) {
            final XSSFSheet oldSheet = oldWorkbook.getSheet("Blank 3-U");

            if(oldSheet != null) {
            final XSSFSheet newSheet = newWorkbook.createSheet(oldSheet.getSheetName());

            newSheet.setDefaultRowHeight(oldSheet.getDefaultRowHeight());
            newSheet.setDefaultColumnWidth(oldSheet.getDefaultColumnWidth());

            // Copy content
            for (int rowNumber = oldSheet.getFirstRowNum(); rowNumber < oldSheet.getLastRowNum(); rowNumber++) {
                final XSSFRow oldRow = oldSheet.getRow(rowNumber);
                if (oldRow != null) {
                    final XSSFRow newRow = newSheet.createRow(rowNumber);
                    newRow.setHeight(oldRow.getHeight());

                    for (int columnNumber = oldRow.getFirstCellNum(); columnNumber < oldRow
                            .getLastCellNum(); columnNumber++) {
                        newSheet.setColumnWidth(columnNumber, oldSheet.getColumnWidth(columnNumber));

                        final XSSFCell oldCell = oldRow.getCell(columnNumber);
                        if (oldCell != null) {
                            final XSSFCell newCell = newRow.createCell(columnNumber);

                            // Copy value
//                            setCellValue(newCell, getCellValue(oldCell));

//                            // Copy style
//                            XSSFCellStyle newCellStyle = newWorkbook.createCellStyle();
//                            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//                            newCell.setCellStyle(newCellStyle);
                        }
                    }
                 }
               }
            }// if Blank 3-U
        //}

        try {
            oldWorkbook.close();
            newWorkbook.write(new FileOutputStream("new.xlsx"));
            newWorkbook.close();
            long nanosecs = (System.currentTimeMillis() - startTime);
            System.out.println("File is prepared and Time taken to create: " + (nanosecs/1000) + " secs" );
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }
    }

//    private static void setCellValue(final XSSFCell cell, final Object value) {
//        if (value instanceof Boolean) {
//            cell.setCellValue((boolean) value);
//        } else if (value instanceof Byte) {
//            cell.setCellValue((byte) value);
//        } else if (value instanceof Double) {
//            cell.setCellValue((double) value);
//        } else if (value instanceof String) {
//            String value1 =  (String) value;
//            cell.setCellValue( (value1.startsWith("=") ? value1.substring(1) : value1 ));
//        } else {
//            throw new IllegalArgumentException();
//        }
//    }

//    private static Object getCellValue(final XSSFCell cell) {
//        switch (cell.getCellType()) {
//            case BOOLEAN:
//                return cell.getBooleanCellValue(); // boolean
//            case ERROR:
//                return cell.getErrorCellValue(); // byte
//            case NUMERIC:
//                return cell.getNumericCellValue(); // double
//            case STRING:
//            case BLANK:
//                return cell.getStringCellValue(); // String
//            case FORMULA:
//                return  "=" + cell.getCellFormula(); // String for formula
//            default:
//                throw new IllegalArgumentException();
//        }
//    }
}