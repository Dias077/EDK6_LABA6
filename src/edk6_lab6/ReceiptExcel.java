package edk6_lab6;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReceiptExcel extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread { // Поток запуска MS Excel

        public void run() {
            
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Текущий катаолог
            try {
                modifData(dir + "receipt_template.xls", dir + "receipt.xls", 
                        jTextField_FIO.getText(),
                        jTextField_Vacancy.getText(), 
                        jTextField_Salary1.getText(), 
                        jTextField_Employment.getText(),
                        jTextField_Adres.getText(),
                        jTextField_Number.getText(),
                        jTextField_Mail.getText(),
                        jTextField_Citizenship.getText(),
                        jTextField_Education.getText(),
                        jTextField_Data.getText(),
                        jTextField_Status.getText(),
                        jTextField_Year.getText(),
                        jTextField_Place.getText(),
                        jTextField_Faculty.getText(),
                        jTextField_Specialization.getText()); // Вызов метода создания отчета
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "receipt.xls").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "receipt.xls")); // Запуск отчета в MS Excel
                }
            } catch (Exception ex) {
                System.err.println("Error modifData!");
                ex.printStackTrace();
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    // Метод создания отчета
    private void modifData(String inputFileName, String outputFileName, String FIO, String vacancy,
            String salary, String employment, String adres, String number,
            String mail, String citizenship, String education, String data,
            String status, String year, String place, String faculty, String specialization) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFileName))); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Первый лист в документе MS Excel
        sheet.getRow(0).getCell(3).setCellValue(FIO);
        sheet.getRow(3).getCell(3).setCellValue(vacancy);
        sheet.getRow(5).getCell(5).setCellValue(salary);
        sheet.getRow(6).getCell(3).setCellValue(employment);
        sheet.getRow(8).getCell(3).setCellValue(adres);
        sheet.getRow(9).getCell(0).setCellValue(number);
        sheet.getRow(9).getCell(8).setCellValue(mail);
        sheet.getRow(14).getCell(6).setCellValue(citizenship);
        sheet.getRow(16).getCell(6).setCellValue(education);
        sheet.getRow(18).getCell(6).setCellValue(data);
        sheet.getRow(20).getCell(6).setCellValue(status);
        sheet.getRow(25).getCell(0).setCellValue(year);
        sheet.getRow(27).getCell(1).setCellValue(place);
        sheet.getRow(28).getCell(2).setCellValue(faculty);
        sheet.getRow(29).getCell(2).setCellValue(specialization);
        
        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public ReceiptExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton1 = new javax.swing.JButton();
        jTextField_FIO = new javax.swing.JTextField();
        jTextField_Vacancy = new javax.swing.JTextField();
        jTextField_Salary1 = new javax.swing.JTextField();
        jTextField_Mail = new javax.swing.JTextField();
        jTextField_Adres = new javax.swing.JTextField();
        jTextField_Employment = new javax.swing.JTextField();
        jTextField_Number = new javax.swing.JTextField();
        jTextField_Citizenship = new javax.swing.JTextField();
        jTextField_Education = new javax.swing.JTextField();
        jTextField_Data = new javax.swing.JTextField();
        jTextField_Status = new javax.swing.JTextField();
        jTextField_Year = new javax.swing.JTextField();
        jTextField_Place = new javax.swing.JTextField();
        jTextField_Faculty = new javax.swing.JTextField();
        jTextField_Specialization = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Работа с Excel");
        setResizable(false);
        getContentPane().setLayout(null);

        jButton1.setText("в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(550, 470, 66, 22);

        jTextField_FIO.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        getContentPane().add(jTextField_FIO);
        jTextField_FIO.setBounds(200, 40, 220, 30);

        jTextField_Vacancy.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Vacancy);
        jTextField_Vacancy.setBounds(220, 80, 180, 24);

        jTextField_Salary1.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Salary1);
        jTextField_Salary1.setBounds(300, 110, 70, 20);

        jTextField_Mail.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Mail);
        jTextField_Mail.setBounds(490, 160, 110, 20);

        jTextField_Adres.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Adres);
        jTextField_Adres.setBounds(250, 150, 140, 30);

        jTextField_Employment.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Employment);
        jTextField_Employment.setBounds(260, 130, 110, 20);

        jTextField_Number.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Number);
        jTextField_Number.setBounds(40, 160, 110, 20);

        jTextField_Citizenship.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Citizenship);
        jTextField_Citizenship.setBounds(460, 280, 140, 20);

        jTextField_Education.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Education);
        jTextField_Education.setBounds(460, 300, 140, 20);

        jTextField_Data.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Data);
        jTextField_Data.setBounds(460, 330, 140, 20);

        jTextField_Status.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Status);
        jTextField_Status.setBounds(460, 350, 140, 20);

        jTextField_Year.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_Year.setToolTipText("");
        getContentPane().add(jTextField_Year);
        jTextField_Year.setBounds(30, 420, 60, 20);

        jTextField_Place.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Place);
        jTextField_Place.setBounds(50, 450, 150, 20);

        jTextField_Faculty.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Faculty);
        jTextField_Faculty.setBounds(60, 470, 150, 20);

        jTextField_Specialization.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Specialization);
        jTextField_Specialization.setBounds(60, 490, 150, 20);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/edk6_lab6/receipt.png"))); // NOI18N
        getContentPane().add(jLabel2);
        jLabel2.setBounds(0, 0, 631, 520);

        setSize(new java.awt.Dimension(648, 556));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton1ActionPerformed

    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        
        
        //</editor-fold>
        

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ReceiptExcel().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JTextField jTextField_Adres;
    private javax.swing.JTextField jTextField_Citizenship;
    private javax.swing.JTextField jTextField_Data;
    private javax.swing.JTextField jTextField_Education;
    private javax.swing.JTextField jTextField_Employment;
    private javax.swing.JTextField jTextField_FIO;
    private javax.swing.JTextField jTextField_Faculty;
    private javax.swing.JTextField jTextField_Mail;
    private javax.swing.JTextField jTextField_Number;
    private javax.swing.JTextField jTextField_Place;
    private javax.swing.JTextField jTextField_Salary1;
    private javax.swing.JTextField jTextField_Specialization;
    private javax.swing.JTextField jTextField_Status;
    private javax.swing.JTextField jTextField_Vacancy;
    private javax.swing.JTextField jTextField_Year;
    // End of variables declaration//GEN-END:variables
}
