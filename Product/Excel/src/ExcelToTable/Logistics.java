
package ExcelToTable;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.FileOutputStream;

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
import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.*;
import java.util.logging.*;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Logistics extends javax.swing.JFrame {


    public Logistics() {
        initComponents();
    }

    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jToggleButton1 = new javax.swing.JToggleButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        jEditorPane1 = new javax.swing.JEditorPane();
        jPasswordField1 = new javax.swing.JPasswordField();
        jMenuBar2 = new javax.swing.JMenuBar();
        jMenu3 = new javax.swing.JMenu();
        jMenu4 = new javax.swing.JMenu();
        jScrollPane2 = new javax.swing.JScrollPane();
        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jPanel4 = new javax.swing.JPanel();
        jPanel5 = new javax.swing.JPanel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jTextField2 = new javax.swing.JTextField();
        jTextField3 = new javax.swing.JTextField();
        jTextField4 = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jTextField5 = new javax.swing.JTextField();
        jLabel15 = new javax.swing.JLabel();
        jLabel16 = new javax.swing.JLabel();
        jTextField10 = new javax.swing.JTextField();
        jTextField12 = new javax.swing.JTextField();
        jComboBox2 = new javax.swing.JComboBox<>();
        jButton2 = new javax.swing.JButton();
        jPanel6 = new javax.swing.JPanel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jTextField6 = new javax.swing.JTextField();
        jTextField7 = new javax.swing.JTextField();
        jTextField8 = new javax.swing.JTextField();
        jTextField9 = new javax.swing.JTextField();
        jLabel14 = new javax.swing.JLabel();
        jButton3 = new javax.swing.JButton();
        jButton4 = new javax.swing.JButton();
        jButton5 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jButton7 = new javax.swing.JButton();
        jButton8 = new javax.swing.JButton();

        jToggleButton1.setText("jToggleButton1");

        jScrollPane1.setViewportView(jEditorPane1);

        jPasswordField1.setText("jPasswordField1");

        jMenu3.setText("File");
        jMenuBar2.add(jMenu3);

        jMenu4.setText("Edit");
        jMenuBar2.add(jMenu4);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setPreferredSize(new java.awt.Dimension(810, 771));

        jPanel1.setBackground(new java.awt.Color(40, 67, 135));

        jLabel1.setFont(new java.awt.Font("Andale Mono", 0, 18)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setText("Logistics");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(366, 366, 366)
                .addComponent(jLabel1)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jButton1.setBackground(new java.awt.Color(40, 67, 135));
        jButton1.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton1.setForeground(new java.awt.Color(255, 255, 255));
        jButton1.setText("Import New Order");
        jButton1.setBorder(null);
        jButton1.setBorderPainted(false);
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jPanel3.setBackground(new java.awt.Color(223, 228, 253));

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );

        jPanel4.setBackground(new java.awt.Color(223, 228, 253));

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );

        jPanel5.setBackground(new java.awt.Color(223, 228, 253));

        jLabel6.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel6.setForeground(new java.awt.Color(120, 120, 120));
        jLabel6.setText("Number");

        jLabel7.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel7.setForeground(new java.awt.Color(120, 120, 120));
        jLabel7.setText("Status");

        jLabel8.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel8.setForeground(new java.awt.Color(120, 120, 120));
        jLabel8.setText("Origin - Destination");

        jLabel9.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel9.setForeground(new java.awt.Color(120, 120, 120));
        jLabel9.setText("Date");

        jLabel2.setFont(new java.awt.Font("Andale Mono", 1, 16)); // NOI18N
        jLabel2.setText("Order Details");

        jTextField1.setBackground(new java.awt.Color(223, 228, 253));
        jTextField1.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField1.setBorder(null);
        jTextField1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField1ActionPerformed(evt);
            }
        });

        jTextField2.setBackground(new java.awt.Color(223, 228, 253));
        jTextField2.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField2.setBorder(null);
        jTextField2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField2ActionPerformed(evt);
            }
        });

        jTextField3.setBackground(new java.awt.Color(223, 228, 253));
        jTextField3.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField3.setBorder(null);
        jTextField3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField3ActionPerformed(evt);
            }
        });

        jTextField4.setBackground(new java.awt.Color(223, 228, 253));
        jTextField4.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField4.setBorder(null);
        jTextField4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField4ActionPerformed(evt);
            }
        });

        jLabel3.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel3.setText("-");

        jTextField5.setBackground(new java.awt.Color(223, 228, 253));
        jTextField5.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField5.setBorder(null);
        jTextField5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField5ActionPerformed(evt);
            }
        });

        jLabel15.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel15.setForeground(new java.awt.Color(120, 120, 120));
        jLabel15.setText("Quantity");

        jLabel16.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel16.setForeground(new java.awt.Color(120, 120, 120));
        jLabel16.setText("Route Distance");

        jTextField10.setBackground(new java.awt.Color(223, 228, 253));
        jTextField10.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField10.setBorder(null);
        jTextField10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField10ActionPerformed(evt);
            }
        });

        jTextField12.setBackground(new java.awt.Color(223, 228, 253));
        jTextField12.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField12.setBorder(null);
        jTextField12.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField12ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addGap(29, 29, 29)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel6)
                    .addComponent(jLabel2)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel5Layout.createSequentialGroup()
                                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel8)
                                    .addGroup(jPanel5Layout.createSequentialGroup()
                                        .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, 132, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jLabel3)))
                                .addGap(26, 26, 26)
                                .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jLabel15)
                            .addComponent(jTextField12))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jTextField10)
                            .addGroup(jPanel5Layout.createSequentialGroup()
                                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel16)
                                    .addComponent(jLabel9)
                                    .addComponent(jLabel7)
                                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, 329, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, 329, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGap(0, 0, Short.MAX_VALUE)))))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addGap(24, 24, 24)
                .addComponent(jLabel2)
                .addGap(18, 18, 18)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(jLabel7))
                .addGap(18, 18, 18)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel8)
                    .addComponent(jLabel9))
                .addGap(18, 18, 18)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3)
                    .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel15)
                    .addComponent(jLabel16))
                .addGap(18, 18, 18)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField10, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField12, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(45, Short.MAX_VALUE))
        );

        jComboBox2.setBackground(new java.awt.Color(40, 67, 135));
        jComboBox2.setForeground(new java.awt.Color(255, 255, 255));
        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel<>(new String[]{ "Select Order", "Order 1", "Order 2", "Order 3","Order 4","Order 5" }));
        jComboBox2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox2ActionPerformed(evt);
            }
        });

        jButton2.setBackground(new java.awt.Color(40, 67, 135));
        jButton2.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton2.setForeground(new java.awt.Color(255, 255, 255));
        jButton2.setText("View Employee Details");
        jButton2.setBorder(null);
        jButton2.setBorderPainted(false);
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jPanel6.setBackground(new java.awt.Color(223, 228, 253));

        jLabel10.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel10.setForeground(new java.awt.Color(120, 120, 120));
        jLabel10.setText("ID");

        jLabel11.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel11.setForeground(new java.awt.Color(120, 120, 120));
        jLabel11.setText("Name");

        jLabel12.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel12.setForeground(new java.awt.Color(120, 120, 120));
        jLabel12.setText("Role");

        jLabel13.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel13.setForeground(new java.awt.Color(120, 120, 120));

        jLabel4.setFont(new java.awt.Font("Andale Mono", 1, 16)); // NOI18N
        jLabel4.setText("Employee Details");

        jTextField6.setBackground(new java.awt.Color(223, 228, 253));
        jTextField6.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField6.setBorder(null);
        jTextField6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField6ActionPerformed(evt);
            }
        });

        jTextField7.setBackground(new java.awt.Color(223, 228, 253));
        jTextField7.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField7.setBorder(null);
        jTextField7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField7ActionPerformed(evt);
            }
        });

        jTextField8.setBackground(new java.awt.Color(223, 228, 253));
        jTextField8.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField8.setBorder(null);
        jTextField8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField8ActionPerformed(evt);
            }
        });

        jTextField9.setBackground(new java.awt.Color(223, 228, 253));
        jTextField9.setFont(new java.awt.Font("Andale Mono", 0, 15)); // NOI18N
        jTextField9.setBorder(null);
        jTextField9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField9ActionPerformed(evt);
            }
        });

        jLabel14.setFont(new java.awt.Font("Andale Mono", 0, 13)); // NOI18N
        jLabel14.setForeground(new java.awt.Color(120, 120, 120));
        jLabel14.setText("Vehicle Number");

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addGap(29, 29, 29)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel10)
                    .addComponent(jLabel4)
                    .addGroup(jPanel6Layout.createSequentialGroup()
                        .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel12)
                            .addComponent(jTextField8, javax.swing.GroupLayout.PREFERRED_SIZE, 132, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(253, 253, 253)
                        .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel11)
                            .addComponent(jTextField7, javax.swing.GroupLayout.PREFERRED_SIZE, 329, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addComponent(jLabel14)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jLabel13))
                            .addComponent(jTextField9, javax.swing.GroupLayout.PREFERRED_SIZE, 195, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel6Layout.setVerticalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addGap(24, 24, 24)
                .addComponent(jLabel4)
                .addGap(18, 18, 18)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel10)
                    .addComponent(jLabel11))
                .addGap(18, 18, 18)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField7, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel12)
                    .addComponent(jLabel13)
                    .addComponent(jLabel14))
                .addGap(18, 18, 18)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField8, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField9, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jButton3.setBackground(new java.awt.Color(40, 67, 135));
        jButton3.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton3.setForeground(new java.awt.Color(255, 255, 255));
        jButton3.setText("Logout");
        jButton3.setBorder(null);
        jButton3.setBorderPainted(false);
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jButton4.setBackground(new java.awt.Color(40, 67, 135));
        jButton4.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton4.setForeground(new java.awt.Color(255, 255, 255));
        jButton4.setText("View Max Quantity Order");
        jButton4.setBorder(null);
        jButton4.setBorderPainted(false);
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        jButton5.setBackground(new java.awt.Color(40, 67, 135));
        jButton5.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton5.setForeground(new java.awt.Color(255, 255, 255));
        jButton5.setText("View Order with Longest Route Distance");
        jButton5.setBorder(null);
        jButton5.setBorderPainted(false);
        jButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton5ActionPerformed(evt);
            }
        });

        jButton6.setBackground(new java.awt.Color(40, 67, 135));
        jButton6.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton6.setForeground(new java.awt.Color(255, 255, 255));
        jButton6.setText("View Min Quantity Order");
        jButton6.setBorder(null);
        jButton6.setBorderPainted(false);
        jButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton6ActionPerformed(evt);
            }
        });

        jButton7.setBackground(new java.awt.Color(40, 67, 135));
        jButton7.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton7.setForeground(new java.awt.Color(255, 255, 255));
        jButton7.setText("View Order with Smallest Route Distance");
        jButton7.setBorder(null);
        jButton7.setBorderPainted(false);
        jButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton7ActionPerformed(evt);
            }
        });

        jButton8.setBackground(new java.awt.Color(40, 67, 135));
        jButton8.setFont(new java.awt.Font("Andale Mono", 0, 14)); // NOI18N
        jButton8.setForeground(new java.awt.Color(255, 255, 255));
        jButton8.setText("Print");
        jButton8.setBorder(null);
        jButton8.setBorderPainted(false);
        jButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton8ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addGap(30, 30, 30)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel5, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jButton4, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(70, 70, 70)
                        .addComponent(jButton6, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                        .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 380, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 353, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                        .addComponent(jButton5, javax.swing.GroupLayout.PREFERRED_SIZE, 360, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jButton7, javax.swing.GroupLayout.PREFERRED_SIZE, 340, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jButton8, javax.swing.GroupLayout.PREFERRED_SIZE, 269, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 269, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(30, 30, 30)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(layout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(21, 21, 21)
                        .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addContainerGap(734, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(15, 15, 15)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton4, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton6, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jButton5, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton7, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 26, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGap(18, 18, 18)
                        .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton8, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(27, 27, 27))))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        AddDataToTable_1 obj = new AddDataToTable_1();
                obj.setVisible(true);
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jComboBox2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox2ActionPerformed
        int selectedIndex = jComboBox2.getSelectedIndex();
        System.out.println(selectedIndex);
        try {  
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    PreparedStatement ps;
                    ResultSet rs;
                    String query = "SELECT * FROM orderdetails LIMIT ?, 1;";
                    ps = MyConnection.getConnection().prepareStatement(query);

                    ps.setInt(1,selectedIndex-1);
                    
                    rs = ps.executeQuery();

                    if (rs.next()) {
                        
                        String data1 = rs.getString("Order ID");
                        jTextField1.setText(data1);
                        String data2 = rs.getString("Status");
                        jTextField2.setText(data2);
                        String data3 = rs.getString("Origin");
                        jTextField3.setText(data3);
                        String data4 = rs.getString("Destination");
                        jTextField4.setText(data4);
                        String data5 = rs.getString("Date");
                        jTextField5.setText(data5);
                        String data6 = rs.getString("Quantity");
                        jTextField12.setText(data6);
                        String data7 = rs.getString("RouteDistance");
                        jTextField10.setText(data7);    
                    }       
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
    }//GEN-LAST:event_jComboBox2ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        // TODO add your handling code here:
        // set the value of the text field based on the selected index
                int selectedIndex = jComboBox2.getSelectedIndex();
                
                try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    
                    PreparedStatement ps;
                    ResultSet rs;
                    PreparedStatement ps2;
                    ResultSet rs2;
                    
                    String query1 = "SELECT * FROM orderdetails LIMIT ?, 1;";
                    ps = MyConnection.getConnection().prepareStatement(query1);
                    ps.setInt(1,selectedIndex-1);
                    rs = ps.executeQuery();
                    
                    if (rs.next()) {
                        
                        int empdata = rs.getInt("Employee ID");
                        System.out.println(empdata);
                        ps2 = MyConnection.getConnection().prepareStatement("SELECT * FROM employeedetails LIMIT ?, 1");
                        ps2.setInt(1,empdata-1);
                        
                        rs2 = ps2.executeQuery();
                        
                        if (rs2.next()) {
                        
                        
                        
                        
                        int data1 = rs2.getInt("ID");
                        jTextField6.setText(String.valueOf(data1));
                        String data2 = rs2.getString("Name");
                        jTextField7.setText(data2);
                        String data3 = rs2.getString("Role");
                        jTextField8.setText(data3);
                        String data4 = rs2.getString("Vehicle Number");
                        jTextField9.setText(data4);}
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }       
            
   
        
        
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jTextField2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField2ActionPerformed

    private void jTextField3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField3ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField3ActionPerformed

    private void jTextField4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField4ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField4ActionPerformed

    private void jTextField5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField5ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField5ActionPerformed

    private void jTextField1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField1ActionPerformed

    private void jTextField6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField6ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField6ActionPerformed

    private void jTextField7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField7ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField7ActionPerformed

    private void jTextField8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField8ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField8ActionPerformed

    private void jTextField9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField9ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField9ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        // TODO add your handling code here:
        NewJFrame obj = new NewJFrame();
                obj.setVisible(true);
                
                this.dispose();
                
    }//GEN-LAST:event_jButton3ActionPerformed

    private void jTextField10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField10ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField10ActionPerformed

    private void jTextField12ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField12ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField12ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        // TODO add your handling code here:
    try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    PreparedStatement ps;
                    ResultSet rs;
                    String query = "SELECT * FROM orderdetails ORDER BY Quantity DESC LIMIT 0,1;";
                    //SELECT * FROM orderdetails ORDER BY quanity ASC LIMIT 0,1;

                    ps = MyConnection.getConnection().prepareStatement(query);

                    
                    
                    rs = ps.executeQuery();

                    // retrieve the data and display it in the text field
                    if (rs.next()) {
                        String data1 = rs.getString("Order ID");
                        jTextField1.setText(data1);
                        String data2 = rs.getString("Status");
                        jTextField2.setText(data2);
                        String data3 = rs.getString("Origin");
                        jTextField3.setText(data3);
                        String data4 = rs.getString("Destination");
                        jTextField4.setText(data4);
                        String data5 = rs.getString("Date");
                        jTextField5.setText(data5);
                        String data6 = rs.getString("Quantity");
                        jTextField12.setText(data6);
                        String data7 = rs.getString("RouteDistance");
                        jTextField10.setText(data7);
                        
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    
                    PreparedStatement ps;
                    ResultSet rs;
                    PreparedStatement ps2;
                    ResultSet rs2;
                    
                    String query1 = "SELECT * FROM orderdetails ORDER BY Quantity DESC LIMIT 0,1;";
                    ps = MyConnection.getConnection().prepareStatement(query1);
                    //ps.setInt(1,selectedIndex-1);
                    rs = ps.executeQuery();
                    
                    if (rs.next()) {
                        
                        int empdata = rs.getInt("Employee ID");
                        System.out.println(empdata);
                        ps2 = MyConnection.getConnection().prepareStatement("SELECT * FROM employeedetails LIMIT ?, 1");
                        ps2.setInt(1,empdata-1);
                        
                        rs2 = ps2.executeQuery();
                        
                        if (rs2.next()) {
                        
                        
                        
                        
                        int data1 = rs2.getInt("ID");
                        jTextField6.setText(String.valueOf(data1));
                        String data2 = rs2.getString("Name");
                        jTextField7.setText(data2);
                        String data3 = rs2.getString("Role");
                        jTextField8.setText(data3);
                        String data4 = rs2.getString("Vehicle Number");
                        jTextField9.setText(data4);}
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }    
    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
        // TODO add your handling code here:
        try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    PreparedStatement ps;
                    ResultSet rs;
                    String query = "SELECT * FROM orderdetails ORDER BY RouteDistance DESC LIMIT 0,1;";
                    //SELECT * FROM orderdetails ORDER BY quanity ASC LIMIT 0,1;

                    ps = MyConnection.getConnection().prepareStatement(query);

                    
                    
                    rs = ps.executeQuery();

                    // retrieve the data and display it in the text field
                    if (rs.next()) {
                        String data1 = rs.getString("Order ID");
                        jTextField1.setText(data1);
                        String data2 = rs.getString("Status");
                        jTextField2.setText(data2);
                        String data3 = rs.getString("Origin");
                        jTextField3.setText(data3);
                        String data4 = rs.getString("Destination");
                        jTextField4.setText(data4);
                        String data5 = rs.getString("Date");
                        jTextField5.setText(data5);
                        String data6 = rs.getString("Quantity");
                        jTextField12.setText(data6);
                        String data7 = rs.getString("RouteDistance");
                        jTextField10.setText(data7);
                        
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
        try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    
                    PreparedStatement ps;
                    ResultSet rs;
                    PreparedStatement ps2;
                    ResultSet rs2;
                    
                    String query1 = "SELECT * FROM orderdetails ORDER BY RouteDistance DESC LIMIT 0,1;";
                    ps = MyConnection.getConnection().prepareStatement(query1);
                    //ps.setInt(1,selectedIndex-1);
                    rs = ps.executeQuery();
                    
                    if (rs.next()) {
                        
                        int empdata = rs.getInt("Employee ID");
                        System.out.println(empdata);
                        ps2 = MyConnection.getConnection().prepareStatement("SELECT * FROM employeedetails LIMIT ?, 1");
                        ps2.setInt(1,empdata-1);
                        
                        rs2 = ps2.executeQuery();
                        
                        if (rs2.next()) {
                        
                        
                        
                        
                        int data1 = rs2.getInt("ID");
                        jTextField6.setText(String.valueOf(data1));
                        String data2 = rs2.getString("Name");
                        jTextField7.setText(data2);
                        String data3 = rs2.getString("Role");
                        jTextField8.setText(data3);
                        String data4 = rs2.getString("Vehicle Number");
                        jTextField9.setText(data4);}
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
    }//GEN-LAST:event_jButton5ActionPerformed

    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton6ActionPerformed
        // TODO add your handling code here:
    try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    PreparedStatement ps;
                    ResultSet rs;
                    String query = "SELECT * FROM orderdetails ORDER BY Quantity ASC LIMIT 0,1;";
                    //SELECT * FROM orderdetails ORDER BY quanity ASC LIMIT 0,1;

                    ps = MyConnection.getConnection().prepareStatement(query);

                    
                    
                    rs = ps.executeQuery();

                    // retrieve the data and display it in the text field
                    if (rs.next()) {
                        String data1 = rs.getString("Order ID");
                        jTextField1.setText(data1);
                        String data2 = rs.getString("Status");
                        jTextField2.setText(data2);
                        String data3 = rs.getString("Origin");
                        jTextField3.setText(data3);
                        String data4 = rs.getString("Destination");
                        jTextField4.setText(data4);
                        String data5 = rs.getString("Date");
                        jTextField5.setText(data5);
                        String data6 = rs.getString("Quantity");
                        jTextField12.setText(data6);
                        String data7 = rs.getString("RouteDistance");
                        jTextField10.setText(data7);
                        
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }   
    try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    
                    PreparedStatement ps;
                    ResultSet rs;
                    PreparedStatement ps2;
                    ResultSet rs2;
                    
                    String query1 = "SELECT * FROM orderdetails ORDER BY Quantity ASC LIMIT 0,1;";
                    ps = MyConnection.getConnection().prepareStatement(query1);
                    //ps.setInt(1,selectedIndex-1);
                    rs = ps.executeQuery();
                    
                    if (rs.next()) {
                        
                        int empdata = rs.getInt("Employee ID");
                        System.out.println(empdata);
                        ps2 = MyConnection.getConnection().prepareStatement("SELECT * FROM employeedetails LIMIT ?, 1");
                        ps2.setInt(1,empdata-1);
                        
                        rs2 = ps2.executeQuery();
                        
                        if (rs2.next()) {
                        
                        
                        
                        
                        int data1 = rs2.getInt("ID");
                        jTextField6.setText(String.valueOf(data1));
                        String data2 = rs2.getString("Name");
                        jTextField7.setText(data2);
                        String data3 = rs2.getString("Role");
                        jTextField8.setText(data3);
                        String data4 = rs2.getString("Vehicle Number");
                        jTextField9.setText(data4);}
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
    }//GEN-LAST:event_jButton6ActionPerformed

    private void jButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton7ActionPerformed
        // TODO add your handling code here:
        try {

                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    PreparedStatement ps;
                    ResultSet rs;
                    String query = "SELECT * FROM orderdetails ORDER BY RouteDistance ASC LIMIT 0,1;";
                    ps = MyConnection.getConnection().prepareStatement(query);
                    rs = ps.executeQuery();
                    // retrieve the data and display it in the text field
                    if (rs.next()) {
                        String data1 = rs.getString("Order ID");
                        jTextField1.setText(data1);
                        String data2 = rs.getString("Status");
                        jTextField2.setText(data2);
                        String data3 = rs.getString("Origin");
                        jTextField3.setText(data3);
                        String data4 = rs.getString("Destination");
                        jTextField4.setText(data4);
                        String data5 = rs.getString("Date");
                        jTextField5.setText(data5);
                        String data6 = rs.getString("Quantity");
                        jTextField12.setText(data6);
                        String data7 = rs.getString("RouteDistance");
                        jTextField10.setText(data7);
  
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
        try {
                    
                    
                    // create a connection to the database
                    Class.forName("com.mysql.cj.jdbc.Driver");
                    Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/new_schema", "root", "root");
                    
                    PreparedStatement ps;
                    ResultSet rs;
                    PreparedStatement ps2;
                    ResultSet rs2;
                    
                    String query1 = "SELECT * FROM orderdetails ORDER BY RouteDistance ASC LIMIT 0,1;";
                    ps = MyConnection.getConnection().prepareStatement(query1);
                    //ps.setInt(1,selectedIndex-1);
                    rs = ps.executeQuery();
                    
                    if (rs.next()) {
                        
                        int empdata = rs.getInt("Employee ID");
                        System.out.println(empdata);
                        ps2 = MyConnection.getConnection().prepareStatement("SELECT * FROM employeedetails LIMIT ?, 1");
                        ps2.setInt(1,empdata-1);
                        
                        rs2 = ps2.executeQuery();
                        
                        if (rs2.next()) {
                        
                        
                        
                        
                        int data1 = rs2.getInt("ID");
                        jTextField6.setText(String.valueOf(data1));
                        String data2 = rs2.getString("Name");
                        jTextField7.setText(data2);
                        String data3 = rs2.getString("Role");
                        jTextField8.setText(data3);
                        String data4 = rs2.getString("Vehicle Number");
                        jTextField9.setText(data4);}
                        
                    }
                    
                } catch (ClassNotFoundException | SQLException ex) {
                    ex.printStackTrace();
                }
    }//GEN-LAST:event_jButton7ActionPerformed

    private void jButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton8ActionPerformed
        // TODO add your handling code here:
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
                excelJTableExporter = new XSSFWorkbook();
                XSSFSheet excelSheet = excelJTableExporter.createSheet("JTable Sheet");

                        XSSFRow excelRow = excelSheet.createRow(0);
                        XSSFCell excelCell = excelRow.createCell(0);
                        
                        excelCell.setCellValue("Order ID");
                      
                        XSSFRow excelRow2 = excelSheet.createRow(1);
                        XSSFCell excelCell2 = excelRow2.createCell(0);
                        
                        excelCell2.setCellValue(jTextField1.getText());
                        XSSFCell excelCell3 = excelRow.createCell(1);
                        
                        excelCell3.setCellValue("Status");
                        XSSFCell excelCell4 = excelRow2.createCell(1);
                        
                        
                        excelCell4.setCellValue(jTextField2.getText());
                        XSSFCell excelCell5 = excelRow.createCell(2);
                        
                        excelCell5.setCellValue("Origin");
                        XSSFCell excelCell6 = excelRow2.createCell(2);
                        
                        excelCell6.setCellValue(jTextField3.getText());
                        XSSFCell excelCell7 = excelRow.createCell(3);
                        
                        excelCell7.setCellValue("Destination");
                        XSSFCell excelCell8 = excelRow2.createCell(3);
                        
                        excelCell8.setCellValue(jTextField4.getText());
                        XSSFCell excelCell9 = excelRow.createCell(4);
                        
                        excelCell9.setCellValue("Date");
                        XSSFCell excelCell10 = excelRow2.createCell(4);
                        
                        excelCell10.setCellValue(jTextField5.getText());
                        XSSFCell excelCell11 = excelRow.createCell(5);
                        
                        excelCell11.setCellValue("Quantity");
                        
                        //XSSFRow excelRow12 = excelSheet.createRow(1);
                        XSSFCell excelCell12 = excelRow2.createCell(5);
                        
                        excelCell12.setCellValue(jTextField12.getText());
                        
                        //XSSFRow excelRow13 = excelSheet.createRow(0);
                        XSSFCell excelCell13 = excelRow.createCell(6);
                        
                        excelCell13.setCellValue("Route Distance");
                        
                        //XSSFRow excelRow14 = excelSheet.createRow(1);
                        XSSFCell excelCell14= excelRow2.createCell(6);
                        
                        excelCell14.setCellValue(jTextField10.getText());
                        
                        
                        
                        
                        XSSFRow excelRow3 = excelSheet.createRow(3);
                        XSSFCell excelCell15 = excelRow3.createCell(0);
                        
                        excelCell15.setCellValue("Employee ID");
                      
                        XSSFRow excelRow4 = excelSheet.createRow(4);
                        XSSFCell excelCell16 = excelRow4.createCell(0);
                        
                        excelCell16.setCellValue(jTextField6.getText());
                        
                        XSSFCell excelCell17 = excelRow3.createCell(1);
                        
                        excelCell17.setCellValue("Name");
                        
                        XSSFCell excelCell18 = excelRow4.createCell(1);
                        
                        excelCell18.setCellValue(jTextField7.getText());
                        
                        XSSFCell excelCell19 = excelRow3.createCell(2);
                        
                        excelCell19.setCellValue("Role");
                        
                        XSSFCell excelCell20 = excelRow4.createCell(2);
                        
                        excelCell20.setCellValue(jTextField8.getText());
                        
                        XSSFCell excelCell21 = excelRow3.createCell(3);
                        
                        excelCell21.setCellValue("Vehicle Number");
                        
                        XSSFCell excelCell22 = excelRow4.createCell(3);
                        
                        excelCell22.setCellValue(jTextField9.getText());
                      

                excelFOU = new FileOutputStream(excelFileChooser.getSelectedFile() + ".xlsx");
                excelBOU = new BufferedOutputStream(excelFOU);
                excelJTableExporter.write(excelBOU);
                JOptionPane.showMessageDialog(null, "Exported Successfully");
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
            
                
    }//GEN-LAST:event_jButton8ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Logistics.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Logistics.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Logistics.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Logistics.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Logistics().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButton7;
    private javax.swing.JButton jButton8;
    private javax.swing.JComboBox<String> jComboBox2;
    private javax.swing.JEditorPane jEditorPane1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenu jMenu4;
    private javax.swing.JMenuBar jMenuBar2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel6;
    private javax.swing.JPasswordField jPasswordField1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField10;
    private javax.swing.JTextField jTextField12;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField3;
    private javax.swing.JTextField jTextField4;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    private javax.swing.JTextField jTextField7;
    private javax.swing.JTextField jTextField8;
    private javax.swing.JTextField jTextField9;
    private javax.swing.JToggleButton jToggleButton1;
    // End of variables declaration//GEN-END:variables
}
