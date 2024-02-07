import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteDataToXlsInvoiceEdit {
    public static void main(String[] args) throws IOException {
        
        try {

            //creating input stream using file
            FileInputStream inputStream = new FileInputStream(new File("/Users/Stephen/Documents/hhrestorations_invoice_template.xlsx"));

            //declaring workbook object as our file declared as inputStream
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0); //Sheet object sheet is declared as the first sheet (index 0) of the workbook.

            Object[][] invoice = //creating an array (multiple "arrays", 2 ) of varying data (String, int)
                    {
                            {"Windshield", 1900},
                            {"Water coolant", 40},
                            {"Oil Change", 70},
                            {"Tires", 200},
                            {"Windshield", 1900},
                            {"Water coolant", 40},
                            {"Oil Change", 70},
                            {"Tires", 200}
                    };

                //create Cell object
                //14 is row in which first service description appears
                //specify in InvoicerMain
                int rowNum = 14;
                //start Cell at row 15, column C
                Row row = sheet.getRow(rowNum);
                Cell cell = row.getCell(1);

                //cursor cell
                Row rowCursor = row;
                Cell cellCursor = cell;

                //iterate through invoice
                for (int miniArray = 0; miniArray < invoice.length; miniArray++) {
                    //iterate through serviceArray
                    for (int data = 0; data < invoice[miniArray].length; data++) {
                        /* Testing array contents as the program iterates through it
                        System.out.println("miniArray = " + miniArray + "\n" +
                                            "data = " + data + "\n" +
                                            invoice[miniArray][data]);

                         */
                        if (miniArray != 0 && data == 0) {
                            rowNum += 1;
                            row = sheet.getRow(rowNum);
                        }


                        if (data == 0) //required step when iterating through arrays
                        {
                            //ensure that the column is set to the description
                            cell = row.getCell(1);

                            cell.setCellValue((String) invoice[miniArray][data]); //at service Description, set service as string
                        }
                        else {
                            //ensure that the column is set ot the cost
                            cell = row.getCell(2);
                            cell.setCellValue((Integer) invoice[miniArray][data]); //at service cost, set amount as double
                        }
                    }
                }


            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

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
