import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class WriteDataToXls {
    public static void main(String[] args) throws IOException {
        /* //Creating a spreadsheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet1 = workbook.createSheet("Sheet1");

        //Creating rows and columns
        Row r0 = sheet1.createRow(0);
        Cell c0 = r0.createCell(0);
        c0.setCellValue("Hello user!");

        //Create a file
        File f = new File("/Users/Stephen/IdeaProjects/rdwr-excel/TestData.xlsx");
        FileOutputStream fos = new FileOutputStream(f);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("File is written successfully!");
    */

        
        try {

            //creating input stream using file
            FileInputStream inputStream = new FileInputStream(new File("/Users/Stephen/Documents/Brian's Program/TestData.xls"));

            //declaring workbook object as our file declared as inputStream
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0); //Sheet object sheet is declared as the first sheet (index 0) of the workbook.

            //use invoice ArrayList in ExcelIO and serviceArray ArrayList for each element of invoice
            Object[][] invoice = //creating an array (multiple "rows", 2 "columns") of varying data (String, int)
                    {
                            {"Windshield", 1900},
                            {"Water coolant", 40},
                    };

            int rowCount = sheet.getLastRowNum(); //counts how many rows there are by getting index of last row of sheet
            //here, the number of rows should be 0 because it is an "empty" file.
            //System.out.print(rowCount); <-- the statement above is true using this log

            for (Object[] service : invoice) { //for each service in the invoice
                
                //using rowCount (number of rows already present), creating a row at 0, and creating a row at each integer following
                //allows one to create row from last used row in a specified sheet
                Row row = sheet.createRow(++rowCount);

                int columnCount = 0;

                Cell cell = row.createCell(columnCount); //creating a cell at each row
                for (Object field : service) //for each amount in a list (service)
                {
                    //because columnCount = 0, + 1 adds a column to the amount. Therefore, createCell method creates a cell at column index 1.
                    cell = row.createCell(columnCount = columnCount + 1);
                    if (field instanceof String) //if the field is a String
                    {
                        cell.setCellValue((String) field); //set cell value as that String
                    } else if (field instanceof Integer) //if field is an integer
                    {
                        cell.setCellValue((Integer) field); //set cell value as that integer
                    } else if (field instanceof Double) //if field is a double
                    {
                        cell.setCellValue((double) field); //set cell value as that double
                    }
                }

            }

            inputStream.close(); //close input stream

            //set output stream
            FileOutputStream outputStream = new FileOutputStream("/Users/Stephen/Documents/Brian's Program/editTestData.xlsx");

            //write output file
            workbook.write(outputStream);

            //close workbook
            workbook.close();

            //close outpout stream
            outputStream.close();

        }
        catch (IOException | EncryptedDocumentException ex)
        {
            ex.printStackTrace();
        }













    }
}
