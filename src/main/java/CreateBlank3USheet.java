import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class CreateBlank3USheet {

    public static void main(String[] args) {

        // Read xlsx file
        XSSFWorkbook oldWorkbook = null;
        try {
            oldWorkbook = (XSSFWorkbook) WorkbookFactory.create(new File("xlsbtest.xlsb"));
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        final XSSFWorkbook newWorkbook = new XSSFWorkbook();

        // Copy style source
        final StylesTable oldStylesSource = oldWorkbook.getStylesSource();
        final StylesTable newStylesSource = newWorkbook.getStylesSource();
        long startTime = System.currentTimeMillis();

        // Copy sheets
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
                            setCellValue(newCell, getCellValue(oldCell));

//                            // Copy style
                            XSSFCellStyle newCellStyle = newWorkbook.createCellStyle();
                            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                            newCell.setCellStyle(newCellStyle);
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

    private static void setCellValue(final XSSFCell cell, final Object value) {
        if (value instanceof Boolean) {
            cell.setCellValue((boolean) value);
        } else if (value instanceof Byte) {
            cell.setCellValue((byte) value);
        } else if (value instanceof Double) {
            cell.setCellValue((double) value);
        } else if (value instanceof String) {
            cell.setCellValue( (String) value);
        } else {
            throw new IllegalArgumentException();
        }
    }

    private static Object getCellValue(final XSSFCell cell) {
        switch (cell.getCellType()) {
            case BOOLEAN:
                return cell.getBooleanCellValue(); // boolean
            case ERROR:
                return cell.getErrorCellValue(); // byte
            case NUMERIC:
                return cell.getNumericCellValue(); // double
            case STRING:
            case BLANK:
          //  case FORMULA:
                return cell.getStringCellValue(); // String
            case FORMULA:
                return  cell.getRawValue(); // String for formula
            default:
                throw new IllegalArgumentException();
        }
    }
}