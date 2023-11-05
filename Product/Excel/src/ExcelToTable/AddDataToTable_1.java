package ExcelToTable;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.logging.Level;
//import java.util.logging.Logger;
import javax.swing.Action;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPopupMenu;
import javax.swing.JTable;
import javax.swing.KeyStroke;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumn;
import javax.swing.text.DefaultEditorKit;
import javax.swing.text.TextAction;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.IOException;  
//import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.*;  
//import org.apache.poi.ss.usermodel.Sheet;  
//import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
//import java.util.*;
//import java.io.*;
//import org.apache.poi.EncryptedDocumentException;
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
 
/**
 *
 * @author Genuine
 */
public class AddDataToTable_1 extends javax.swing.JFrame {
 
    /**
     * Creates new form AddDataToJTable
     */
    
    public AddDataToTable_1() {
        initComponents();
        addTableHeader();
        
    }
    
    @SuppressWarnings("unchecked")
    private void initComponents() {
 
        
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        

        jButton6 = new javax.swing.JButton();
        jButtonImportExcelToJtable = new javax.swing.JButton();
 
        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
 
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
 
            },
            new String [] {
                "null"
            }
        ));
        jTable1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jTable1MouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(jTable1);
        jTable1.setBackground(new java.awt.Color(223,228,253));
 
        
        jButton6.setText("Export (Excel)");
        jButton6.setBackground(new java.awt.Color(40, 67, 135));
        jButton6.setForeground(new java.awt.Color(255,255, 255));

        jButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton6ActionPerformed(evt);
            }
        });
 
        jButtonImportExcelToJtable.setText("Import (Excel)");
        jButtonImportExcelToJtable.setBackground(new java.awt.Color(40, 67, 135));
        jButtonImportExcelToJtable.setForeground(new java.awt.Color(255, 255, 255));
        jButtonImportExcelToJtable.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonImportExcelToJtableActionPerformed(evt);
            }
        });
 
        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    )
                                .addGap(18, 18, 18)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    ))
                            .addGroup(layout.createSequentialGroup()
                                
                                .addGap(18, 18, 18)
                                ))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 256, Short.MAX_VALUE)
                        )
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 534, Short.MAX_VALUE)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    
                                    )
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                    .addComponent(jButton6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(jButtonImportExcelToJtable, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    
                                    
                                    
                                    )
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                ))
                        .addContainerGap())))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    ))
                            )
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                )
                            .addComponent(jButtonImportExcelToJtable, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton6, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(5, 5, 5)
                                )
                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                ))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    
                    )
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    
                    
                    )
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 251, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(31, Short.MAX_VALUE))
        );
 
        pack();
        setLocationRelativeTo(null);
    }          
 
    DefaultTableModel model;
   
 
    public void addTableHeader() {
        model = (DefaultTableModel) jTable1.getModel();
        Object[] newIdentifiers = new Object[]{"From","To", "Distance"};
        model.setColumnIdentifiers(newIdentifiers);
        
    }
                                           
    private void jTable1MouseClicked(java.awt.event.MouseEvent evt) 
    {                                     
 
    }                                                                                                                                                                                                                                           
    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {                                         
 
        FileOutputStream excelFOU = null;
        BufferedOutputStream excelBOU = null;
        XSSFWorkbook excelJTableExporter = null;
 
        JFileChooser excelFileChooser = new JFileChooser("C:\\Users\\Authentic\\Desktop");
        excelFileChooser.setDialogTitle("Save As");
        FileNameExtensionFilter fnef = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx", "xlsm");
        excelFileChooser.setFileFilter(fnef);
        int excelChooser = excelFileChooser.showSaveDialog(null);
        if (excelChooser == JFileChooser.APPROVE_OPTION) {
 
            try {
                //Import excel poi libraries to netbeans
                excelJTableExporter = new XSSFWorkbook();
                XSSFSheet excelSheet = excelJTableExporter.createSheet("JTable Sheet");
                //            Loop to get jtable columns and rows
                for (int i = 0; i < model.getRowCount(); i++) {
                    XSSFRow excelRow = excelSheet.createRow(i);
                    for (int j = 0; j < model.getColumnCount(); j++) {
                        XSSFCell excelCell = excelRow.createCell(j);
 
                        //Now Get ImageNames From JLabel
                        //get the last column
                        
                        excelCell.setCellValue(model.getValueAt(i, j).toString());
 
//                        Change the values of the fourth column to image paths
                        
                    }
                }   //Append xlsx file extensions to selected files. To create unique file names
                excelFOU = new FileOutputStream(excelFileChooser.getSelectedFile() + ".xlsx");
                excelBOU = new BufferedOutputStream(excelFOU);
                excelJTableExporter.write(excelBOU);
                JOptionPane.showMessageDialog(null, "Exported Successfully!");
            } catch (FileNotFoundException ex) {
                ex.printStackTrace();
            } catch (IOException ex) {
                ex.printStackTrace();
            } finally {
                try {
                    if (excelBOU != null) {
                        excelBOU.close();
                    }
                    if (excelFOU != null) {
                        excelFOU.close();
                    }
                    if (excelJTableExporter != null) {
                        excelJTableExporter.close();
                    }
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
 
        }
        
 
    }      
    File excelFile;
    FileInputStream excelFIS = null;
    BufferedInputStream excelBIS = null;
    XSSFWorkbook excelImportToJTable = null;
 
    private void jButtonImportExcelToJtableActionPerformed(java.awt.event.ActionEvent evt) {                                                           
        double distance = 0.0;
        int rows = 0;
        int columns = 0;
        String place1 = "";
        String place2 = "";
        double sum =0.0;
        String defaultCurrentDirectoryPath = "C:\\Users\\Authentic\\Desktop";
        JFileChooser excelFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
        excelFileChooser.setDialogTitle("Select Excel File");
        FileNameExtensionFilter fnef = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx", "xlsm");
        excelFileChooser.setFileFilter(fnef);
        int excelChooser = excelFileChooser.showOpenDialog(null);
        if (excelChooser == JFileChooser.APPROVE_OPTION) 
        {
            try {
                excelFile = excelFileChooser.getSelectedFile();
                excelFIS = new FileInputStream(excelFile);
                excelBIS = new BufferedInputStream(excelFIS);
                excelImportToJTable = new XSSFWorkbook(excelBIS);
                XSSFSheet excelSheet = excelImportToJTable.getSheetAt(0);
 
                for (int row = 0 ; row <= excelSheet.getLastRowNum(); row++) 
                {
                    //System.out.println(excelSheet.getLastRowNum());
                    XSSFRow excelRow = excelSheet.getRow(row);
                    XSSFCell excelName = excelRow.getCell(0);
                    String value1 = excelName.getStringCellValue();                    
 
                    XSSFRow excelRow2 = excelSheet.getRow(row);
                    XSSFCell excelName2 = excelRow2.getCell(1);
                    String value2 = excelName2.getStringCellValue();
                        //System.out.println(value2);
                        
                    place1 = value1;
                    place2 = value2;
                        
                    System.out.println(place1);
                    System.out.println(place2);
                    //model.addRow(new Object[]{place1,place2, sum});
                        
        if(place1.equalsIgnoreCase("Amble"))
        {
               rows=1;
        }
        else if(place1.equalsIgnoreCase("Ashti"))
        {   
            rows =2;
        }
         else if(place1.equalsIgnoreCase("Brahamgaon"))
        {   
            rows =3;
        }
         else if(place1.equalsIgnoreCase("Dhamari"))
        {   
            rows =4;
        }
         else if(place1.equalsIgnoreCase("Dhirdi"))
        {   
            rows =5;
        }
         else if(place1.equalsIgnoreCase("Gunat"))
        {   
            rows =6;
        }
         else if(place1.equalsIgnoreCase("Hingani D"))
        {   
            rows =7;
        }
         else if(place1.equalsIgnoreCase("Kahnur P"))
        {   
            rows =8;
        }
         else if(place1.equalsIgnoreCase("Kel Sanghavi"))
        {   
            rows =9;
        }
         else if(place1.equalsIgnoreCase("Kerul"))
        {   
            rows =10;
        }
         else if(place1.equalsIgnoreCase("Dlecta"))
        {   
            rows =11;
        }
         else if(place1.equalsIgnoreCase("Koregoan B"))
        {   
            rows =12;
        }
         else if(place1.equalsIgnoreCase("Memanewasti"))
        {   
            rows =13;
        }
         else if(place1.equalsIgnoreCase("Navare"))        
         {   
            rows =14;
        }
         else if(place1.equalsIgnoreCase("Malganga"))
        {   
            rows =15;
        }
         else if(place1.equalsIgnoreCase("Parner"))
        {   
            rows =16;
        }

         else if(place1.equalsIgnoreCase("Ranjani"))
        {   
            rows =17;
        }
         else if(place1.equalsIgnoreCase("Shiral"))
        {   
            rows =18;
        }
         else if(place1.equalsIgnoreCase("Walunj"))
        {   
            rows =19;
        }
         else if(place1.equalsIgnoreCase("Inamgaon"))
        {   
            rows =20;
        }
         else if(place1.equalsIgnoreCase("Nimone"))
        {   
            rows =21;
        }
         else if(place1.equalsIgnoreCase("Dhavlgaon"))
        {   
            rows =22;
        }
         else if(place1.equalsIgnoreCase("Deodaithan"))
        {   
            rows =23;
        }
         else if(place1.equalsIgnoreCase("Talegaon Dhamder"))
        {   
            rows =24;
        }
         else if(place1.equalsIgnoreCase("Gholapwadi"))
        {   
            rows =25;
        }
         
        else
        System.out.println("Invalid1");
       
         
        if(place2.equalsIgnoreCase("Amble"))
        {
               columns=1;
        }
        else if(place2.equalsIgnoreCase("Ashti"))
        {   
            columns =2;
        }
         else if(place2.equalsIgnoreCase("Brahamgaon"))
        {   
            columns =3;
        }
         else if(place2.equalsIgnoreCase("Dhamari"))
        {   
            columns =4;
        }
         else if(place2.equalsIgnoreCase("Dhirdi"))
        {   
            columns =5;
        }
         else if(place2.equalsIgnoreCase("Gunat"))
        {   
            columns =6;
        }
         else if(place2.equalsIgnoreCase("Hingani D"))
        {   
            columns =7;
        }
         else if(place2.equalsIgnoreCase("Kahnur P"))
        {   
            columns =8;
        }
         else if(place2.equalsIgnoreCase("Kel Sanghavi"))
        {   
            columns =9;
        }
         else if(place2.equalsIgnoreCase("Kerul"))
        {   
            columns =10;
        }
         else if(place2.equalsIgnoreCase("Dlecta"))
        {   
            columns =11;
        }
         else if(place2.equalsIgnoreCase("Koregaon B"))
        {   
            columns =12;
        }
         else if(place2.equalsIgnoreCase("Memanewasti"))
        {   
            columns =13;
        }
         else if(place2.equalsIgnoreCase("Navare"))
        {   
            columns =14;
        }
         else if(place2.equalsIgnoreCase("Malganga"))
        {   
            columns =15;
        }
         else if(place2.equalsIgnoreCase("Parner"))
        {   
            columns =16;
        }
         else if(place2.equalsIgnoreCase("Ranjani"))
        {   
            columns =17;
        }
         else if(place2.equalsIgnoreCase("Shiral"))
        {   
            columns =18;
        }
         else if(place2.equalsIgnoreCase("Walunj"))
        {   
            columns =19;
        }
         else if(place2.equalsIgnoreCase("Inamgaon"))
        {   
            columns =20;
        }
         else if(place2.equalsIgnoreCase("Nimone"))
        {   
            columns =21;
        }
         else if(place2.equalsIgnoreCase("Dhavlgaon"))
        {   
            columns =22;
        }
         else if(place2.equalsIgnoreCase("Deodaithan"))
        {   
            columns =23;
        }
         else if(place2.equalsIgnoreCase("Talegaon Dhamder"))
        {   
            columns =24;
        }
         else if(place2.equalsIgnoreCase("Gholapwadi"))
        {   
            columns =25;
        }
        else
        System.out.println("Invalid2");
        
        
        distance = ReadCellData(rows,columns);
        sum = sum + distance;
        System.out.println(distance);
        
        model.addRow(new Object[]{place1,place2,distance});
                    
        /*if(row == excelSheet.getLastRowNum())
        {
            distance = 0; 
            sum = sum + distance;          
            model.addRow(new Object[]{place1,place2, sum}); 
        }*/
                                    
                }
                JOptionPane.showMessageDialog(null, "Imported Successfully");
            } catch (IOException iOException) {
                JOptionPane.showMessageDialog(null, iOException.getMessage());
            } finally {
                try {
                    if (excelFIS != null) {
                        excelFIS.close();
                    }
                    if (excelBIS != null) {
                        excelBIS.close();
                    }
                    if (excelImportToJTable != null) {
                        excelImportToJTable.close();
                    }
                } catch (IOException iOException) {
                    JOptionPane.showMessageDialog(null, iOException.getMessage());
                }
            }
        }
        
    }
    public double ReadCellData(int vRow, int vColumn)  
    {  
        double value=0.0;          //variable for storing the cell value  
        Workbook wb=null;           //initialize Workbook null  
        DataFormatter dataFormatter = new DataFormatter();
        try  
        {  
        //reading data from a file in the form of bytes  
        FileInputStream fis=new FileInputStream("//Users/kshitijkaria/Desktop/Distance Calculator.xlsx");  
        //constructs an XSSFWorkbook object, by buffering the whole stream into the memory  
        wb=new XSSFWorkbook(fis); 
        }  
        catch(FileNotFoundException e)  
        {  
        e.printStackTrace();  
        }  
        catch(IOException e1)  
        {  
        e1.printStackTrace();  
        }  
        Sheet sheet=wb.getSheetAt(0);   //getting the XSSFSheet object at given index  
        Row row=sheet.getRow(vRow); //returns the logical row  
        Cell cell=row.getCell(vColumn); //getting the cell representing the given column  
        value = cell.getNumericCellValue();    //getting cell value  
        //String cellvalue = dataFormatter.formatCellValue(value);
        System.out.println(value);
        return value; 
        //returns the cell value  
    }
    public  String ReadStringCellData(int vRow, int vColumn)  
    {  
        String value="";          //variable for storing the cell value  
        Workbook wb=null;           //initialize Workbook null  
        DataFormatter dataFormatter = new DataFormatter();
        try  
        {  
        //reading data from a file in the form of bytes  
        FileInputStream fis=new FileInputStream(excelFile);  
        //constructs an XSSFWorkbook object, by buffering the whole stream into the memory  
        wb=new XSSFWorkbook(fis); 
        }  
        catch(FileNotFoundException e)  
        {  
        e.printStackTrace();  
        }  
        catch(IOException e1)  
        {  
        e1.printStackTrace();  
        }  
        Sheet sheet=wb.getSheetAt(0);   //getting the XSSFSheet object at given index  
        Row row=sheet.getRow(vRow); //returns the logical row  
        Cell cell=row.getCell(vColumn); //getting the cell representing the given column  
        value = cell.getStringCellValue();    //getting cell value  
        //String cellvalue = dataFormatter.formatCellValue(value);
        return value;               //returns the cell value  
    }
 
    
    public static void main(String args[]) {
        
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(AddDataToTable_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(AddDataToTable_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(AddDataToTable_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(AddDataToTable_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
       
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new AddDataToTable_1().setVisible(true);
            }
        });  
    }
 
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButtonImportExcelToJtable;    
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
}