import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public interface ExcelIO {

    //-----Declarations-----
    ArrayList<ArrayList<Object>> invoice = new ArrayList<>();
    ArrayList<Object> serviceArray = new ArrayList<Object>();


    //-----Methods-----
    //used to add new ArrayList element containing service and cost of service
    abstract void addService(String service, int cost);

    //test method used to print elements at indexes i and j to ensure the function of
    //addService.
    abstract void printInvoice();

    //creates Workbook object and Sheet object using a created FileInputStream at specified filePath String
    abstract void createWorkbookSheet(String filePath);

    //takes ArrayList of ArrayLists as input and writes data to sheet cells
    abstract void writeData();

    //closes inputStream, writes outputStream, closes workbook then outputStream.
    abstract void closeAndWrite();












    /*Test
    public static void main(String[] args) throws IOException {


        addService("Windshield", 1900);
        addService("Water Coolant", 40);
        addService("Oil Change", 70);
        //addService("Tires", 40);

        printInvoice();
         */
    }