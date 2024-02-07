import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Invoicer implements ExcelIO {

    //-----Declarations-----
    static String fileDestination;

    private JPanel mainPanel;
    private JFormattedTextField directoryField;
    private JButton setPathButton;
    private JButton nextButton;

    protected String directoryPath = directoryField.getText();


    //-----Abstract Methods-----
    @Override
    public void addService(String service, int cost) {}

    @Override
    public void printInvoice() {}

    @Override
    public void createWorkbookSheet(String filePath) {

    }

    @Override
    public void writeData() {

    }

    @Override
    public void closeAndWrite() {

    }


    //-----Main-----
    public static void main(String[] args) throws IOException
    {
        JFrame frame = new JFrame("Invoicer");
        frame.setContentPane(new Invoicer().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
        frame.setBounds(0, 0, 420, 420);
    }



    //-----Action Listeners-----
    public Invoicer() {
        //when an action (click) is made on the "Set Path" button, the fileDestination String is set to whatever is in the text box.
        setPathButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fileDestination = directoryField.getText();
                System.out.println("File Destination: " + fileDestination);
            }
        });

        //Next Button click
        nextButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (e.getSource()==nextButton)
                {
                    InvoicerMain mainWindow = new InvoicerMain();
                }
            }
        });
    }
}
