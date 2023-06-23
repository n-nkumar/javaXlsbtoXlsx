
import com.microsoft.azure.functions.ExecutionContext;
import com.microsoft.azure.functions.OutputBinding;
import com.microsoft.azure.functions.annotation.*;
import com.spire.xls.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.EnumSet;

public class CreateXlsxFromXlsb {

    private static String fileNameStr = "Dovamo.xlsb";

    @FunctionName("BlobTrigger")
    @StorageAccount("AzureWebJobsStorage")
    public void BlobTriggerToBlobTest(
            @BlobTrigger(name = "triggerBlob", path = "samples/{name}", dataType = "binary") byte[] triggerBlob,
            @BindingName("name") String fileName,
            @BlobInput(name = "inputBlob", path = "samples/{name}", dataType = "binary") byte[] inputBlob,
            @BlobOutput(name = "outputBlob", path = "output/{name}", dataType = "binary") OutputBinding<byte[]> outputBlob,
            final ExecutionContext context
    ) {
        fileNameStr = fileName;
        context.getLogger().info("Java Blob trigger function BlobTriggerToBlobTest processed a blob.\n Name: " + fileName + "\n Size: " + triggerBlob.length + " Bytes");
        outputBlob.setValue(inputBlob);
    }

        public static void main(String[] args) {

            long startTime = System.currentTimeMillis();
            Workbook workbook = new Workbook();
            //Open excel from a stream
            FileInputStream fileStream = null;
      try {
            fileStream = new FileInputStream(fileNameStr);
            workbook.loadFromStream(fileStream);
            Worksheet sheet = workbook.getWorksheets().get("Blank 3-U");

            Workbook workbook1 = new Workbook();
            Worksheet sheet1 = workbook1.getWorksheets().get(0);
            sheet1.copyFrom(sheet);

            //Find the cells that contain formula "=SUM(A11,A12)"
            CellRange[] ranges = sheet.findAll("=", EnumSet.of(FindType.Formula), EnumSet.of(ExcelFindOptions.None));

            for(CellRange cell : ranges){
               CellRange copyCell =  sheet1.getCellRange(cell.getRow(), cell.getColumn());
               copyCell.setValue(cell.getFormulaValue().toString());
            }

            // Copy worksheet to destination worsheet in another Excel file

            String result1 = "Copy-" + fileNameStr.substring(0,fileNameStr.lastIndexOf(".xlsb")) +".xlsx";

            // Save the destination file
            workbook1.saveToFile(result1, ExcelVersion.Version2013);
            long endTime = System.currentTimeMillis();
            System.out.println("Time for prepartion of excel file: " + ((endTime - startTime)/1000) );
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }


}
