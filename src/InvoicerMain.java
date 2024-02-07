import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.*;
import java.util.ArrayList;

public class InvoicerMain extends Invoicer implements ExcelIO {


    //-----Objects------
    FileInputStream inputStream;

    Workbook workbook;

    Sheet sheet;



    JFrame frame = new JFrame();
    JLabel label = new JLabel("Hello");

    JButton writeButton = new JButton("Write Data");

    InvoicerMain() {
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(420,420);
        frame.setLayout(null);
        frame.setVisible(true);

        addService("Tires", 400);
        addService("Windshield", 275);
        createWorkbookSheet("/Users/Stephen/Documents/hhrestorations_invoice_template.xlsx");
        writeData();
        closeAndWrite();

        label.setBounds(0, 0, 100, 50);
        frame.add(label);

        writeButton.setBounds(0,0,100, 50);
        frame.add(writeButton);
    }









    //-----Methods-----
    //used to add new ArrayList element containing service and cost of service
    @Override
    public void addService(String service, int cost) {
        //clear any previous elements
        serviceArray.clear();

        serviceArray.add(service); //add String service at index 0
        serviceArray.add(cost); //add int cost at index 1

        //add serviceArray as new element (ArrayList object) of invoice
        invoice.add(new ArrayList<>(serviceArray));
    }

    //test method used to print elements at indexes i and j to ensure the function of addService
    @Override
    public void printInvoice() {
        //iterate through 2D ArrayList
        for (int i = 0; i < invoice.size(); i++)
        {
            //Print index of outer 2D ArrayList
            System.out.println("\n\nInvoice Item Index: " + i);
            //iterate through inner 1D ArrayLists
            for (int j = 0; j < invoice.get(i).size(); j++)
            {
                if (j == 0) System.out.println("Service: " + invoice.get(i).get(j));
                else System.out.println("Cost: " + invoice.get(i).get(j));
            }
        }
    }

    //creates Workbook object and Sheet object using a created FileInputStream at specified filePath String
    @Override
    public void createWorkbookSheet(String filePath) {

        try {
            //create input stream using template .xlsx file
            inputStream = new FileInputStream(new File("/Users/Stephen/Documents/hhrestorations_invoice_template.xlsx"));

            //declaring workbook object as our file declared as inputStream
            workbook = WorkbookFactory.create(inputStream);

            //create workbook sheet object
            sheet = workbook.getSheetAt(0);
        }
        catch (EncryptedDocumentException | IOException ex)
        {
            ex.printStackTrace();
        }

    }



    //takes ArrayList of ArrayLists as input and writes data to sheet cells
    @Override
    public void writeData() {
        //create cell object
        //currently 14, first row which service description appears.
        //maybe at some point add a GUI function to select this on template?
        int rowNum = 14;
        Row row = sheet.getRow(rowNum);
        //index column 1, which is the first column service description appears.
        Cell cell = row.getCell(1);

        //iterate through invoice (outer, 2D ArrayList)
        for (int i = 0; i < invoice.size(); i++)
        {
            //iterate through serviceArray elements (inner, 1D ArrayLists)
            for (int j = 0; j < invoice.get(i).size(); j++)
            {

                //increments rowNum after first row is finished to write in following rows
                if (i != 0 && j == 0)
                {
                    rowNum += 1;
                    row = sheet.getRow(rowNum);
                }

                //check if "looking at" service description cell
                if (j == 0) {
                    //ensure column is set to description column
                    cell = row.getCell(1);

                    //set data
                    cell.setCellValue((String) invoice.get(i).get(j));
                }

                else {

                    //ensure that the column is set at the cost
                    cell = row.getCell(2);
                    cell.setCellValue((Integer) invoice.get(i).get(j));
                }

                //end of inner iterator
                }
            //end of outer iterator
            }
        //end of method
        }



    //closes inputStream, writes outputStream, closes workbook then outputStream.
    @Override
    public void closeAndWrite() {
        try {
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
        catch (IOException ex)
        {
            ex.printStackTrace();
        }

    }
}
