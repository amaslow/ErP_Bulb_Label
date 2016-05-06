package imagesplit;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.font.TextAttribute;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.AttributedString;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.awt.Rectangle;
import java.util.Iterator;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.Document;
import java.util.ArrayList;
import java.util.List;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

public class ErP extends javax.swing.JFrame {

    public ErP() {
        initComponents();
        jProgressBar1.setStringPainted(true);
        jItemTextField.requestFocus();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {
        bindingGroup = new org.jdesktop.beansbinding.BindingGroup();

        buttonGroup1 = new javax.swing.ButtonGroup();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu2 = new javax.swing.JMenu();
        jMenu3 = new javax.swing.JMenu();
        jItemStartButton = new javax.swing.JButton();
        jItemLabel = new javax.swing.JLabel();
        jItemTextField = new javax.swing.JTextField();
        jProgressBar1 = new javax.swing.JProgressBar();
        jListTextField = new javax.swing.JTextField();
        jListLabel = new javax.swing.JLabel();
        jItemRadioButton = new javax.swing.JRadioButton();
        jListRadioButton = new javax.swing.JRadioButton();
        jBrowseButton = new javax.swing.JButton();
        jListStartButton = new javax.swing.JButton();
        jRowCounterLabel = new javax.swing.JLabel();
        PDFCheckBox = new javax.swing.JCheckBox();

        jMenu2.setText("File");
        jMenuBar1.add(jMenu2);

        jMenu3.setText("Edit");
        jMenuBar1.add(jMenu3);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("ErP Bulb Labeling");

        jItemStartButton.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jItemStartButton.setText("START");
        jItemStartButton.setEnabled(false);

        org.jdesktop.beansbinding.Binding binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jItemStartButton, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        jItemStartButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jItemStartButtonActionPerformed(evt);
            }
        });

        jItemLabel.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jItemLabel.setText("1 item:");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jItemLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        jItemTextField.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jItemRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jItemTextField, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        jItemTextField.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jItemTextFieldActionPerformed(evt);
            }
        });

        jListTextField.setFont(new java.awt.Font("Tahoma", 1, 8)); // NOI18N

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jListRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jListTextField, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        jListLabel.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jListLabel.setText("List of items:");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jListRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jListLabel, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        buttonGroup1.add(jItemRadioButton);
        jItemRadioButton.setSelected(true);
        jItemRadioButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jItemRadioButtonActionPerformed(evt);
            }
        });

        buttonGroup1.add(jListRadioButton);

        jBrowseButton.setText("Browse");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jListRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jBrowseButton, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        jBrowseButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBrowseButtonActionPerformed(evt);
            }
        });

        jListStartButton.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jListStartButton.setText("START");

        binding = org.jdesktop.beansbinding.Bindings.createAutoBinding(org.jdesktop.beansbinding.AutoBinding.UpdateStrategy.READ_WRITE, jListRadioButton, org.jdesktop.beansbinding.ELProperty.create("${selected}"), jListStartButton, org.jdesktop.beansbinding.BeanProperty.create("enabled"));
        bindingGroup.addBinding(binding);

        jListStartButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jListStartButtonActionPerformed(evt);
            }
        });

        PDFCheckBox.setText(" + PDF");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(29, 29, 29)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jListRadioButton)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jListLabel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jItemRadioButton)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jItemLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 87, javax.swing.GroupLayout.PREFERRED_SIZE))))
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jRowCounterLabel)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addComponent(jProgressBar1, javax.swing.GroupLayout.DEFAULT_SIZE, 215, Short.MAX_VALUE)
                        .addComponent(jListTextField))
                    .addComponent(jItemTextField, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addComponent(jBrowseButton)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jListStartButton))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(PDFCheckBox)
                        .addComponent(jItemStartButton)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(PDFCheckBox)
                .addGap(1, 1, 1)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(9, 9, 9)
                        .addComponent(jItemRadioButton))
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jItemStartButton, javax.swing.GroupLayout.PREFERRED_SIZE, 37, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jItemLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 38, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jItemTextField, javax.swing.GroupLayout.PREFERRED_SIZE, 37, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jListStartButton, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jListTextField, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jListLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jBrowseButton)
                            .addComponent(jListRadioButton, javax.swing.GroupLayout.DEFAULT_SIZE, 35, Short.MAX_VALUE))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jProgressBar1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jRowCounterLabel)))
        );

        jItemTextField.getAccessibleContext().setAccessibleDescription("");
        jItemTextField.getDocument().addDocumentListener(new DocumentListener() {
            public void changedUpdate(DocumentEvent e) {
                changed();
            }
            public void removeUpdate(DocumentEvent e) {
                changed();
            }
            public void insertUpdate(DocumentEvent e) {
                changed();
            }
            public void changed() {
                if (!jItemTextField.getText().equals("") && jItemRadioButton.isSelected()){
                    jItemStartButton.setEnabled(true);
                }
                else {
                    jItemStartButton.setEnabled(false);
                }
            }
        });

        bindingGroup.bind();

        pack();
    }// </editor-fold>//GEN-END:initComponents
    String mainfolder = "G:\\Share Company Wide\\Company Transfer\\ERP classificatie";
    String productContent = "G:\\Product Content\\PRODUCTS\\";

    int rownr = 0;
    String itemNo = null;
    String sap = null;
    String ean = null;
    String wat = null;
    int en_w = 0;
    XSSFCell log = null;
    String destF;
    File output;
    InputStream inPut = null;
    OutputStream outPut = null;

    private int findRow(XSSFSheet sheet, String item) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(item)) {
                        return row.getRowNum();
                    }
                }
            }
        }
        return 0;
    }

    public void getData(String item, int rownr, String itemNo, String sap, String ean, String wat, int en_w, XSSFCell log, String destF) throws IOException {
        String excelname = mainfolder + "\\ERP bulbs.xlsx";
        List<String> noitem = new ArrayList<String>();
        FileInputStream fis = null;
        fis = new FileInputStream(excelname);
        System.out.println(item);

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        rownr = findRow(sheet, item);

        XSSFRow row = sheet.getRow(rownr);

        XSSFCell itemNo1 = row.getCell(0); // get item number
        itemNo = itemNo1.getStringCellValue();

        XSSFCell sap1 = row.getCell(1); // get sap number
        sap = sap1.getStringCellValue();
        sap = sap.replace(".", "");

        XSSFCell ean1 = row.getCell(2); // get ean code
        ean = ean1.getStringCellValue();

        XSSFCell wat1 = row.getCell(3); // get wattage
        
        //wat = wat1.toString();
        wat = String.valueOf((int) wat1.getNumericCellValue())+"  ";
        //int wat2 = (int) wat1.getNumericCellValue();
        //wat = wat.substring(0, wat.length() - 2);

        XSSFCell en = row.getCell(4); // get class
        String en_1 = en.getStringCellValue();
        en_w = 0;

        switch (en_1) {
            case "A++":
                en_w = 1;
                break;
            case "A+":
                en_w = 2;
                break;
            case "A":
                en_w = 3;
                break;
            case "B":
                en_w = 4;
                break;
            case "C":
                en_w = 5;
                break;
            case "D":
                en_w = 6;
                break;
            case "E":
                en_w = 7;
                break;
            default:
                en_w = 0;
        }

        log = row.getCell(5, org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK); // get logo

        switch (log.getCellType()) {
            case XSSFCell.CELL_TYPE_BLANK:
                log.setCellValue("Logo_0");
                break;
        }

        if (destF.substring(destF.length() - 4).equals("null")) {
            destF = productContent + "\\" + sap;
        }

        if (rownr != 0) {

            label(item, destF, noitem, rownr, itemNo, sap, ean, wat, en_w, log);

            if (jItemRadioButton.isSelected()) {
                File subdir = new File(destF + "\\");
                Desktop desktop = Desktop.getDesktop();
                desktop.open(subdir);
            }
        } else {
            JOptionPane.showMessageDialog(null, "There is no data for item: " + item);
        }
    }

    public void label(String item, String destF, List<String> noitem, int rownr, String itemNo, String sap, String ean, String wat, int en_w, XSSFCell log) throws IOException {
        File sources = new File(mainfolder + "\\ERP_Elements");
        String[] size = {"S", "W"};
        String[] color = {"_BW", ""};
        jProgressBar1.setValue(0);

        if (rownr != 0 && en_w > 0) {
            for (int s = 0; s < size.length; s++) {
                for (int c = 0; c < color.length; c++) {

//                    File subdir = new File(dest.getSelectedFile() + "\\" + sap + "\\");
                    File subdir = new File(destF + "\\");
                    switch (c) {
                        case 0:
                            output = new File(subdir + "\\Energylabel_" + sap + "_" + size[s] + color[c] + ".png");
                            break;
                        case 1:
                            output = new File(subdir + "\\Energylabel_" + sap + "_" + size[s] + "_C.png");
                            break;
                    }
                    BufferedImage base = ImageIO.read(new File(sources + "\\Energy_" + size[s] + "_" + en_w + color[c] + ".png"));
                    int w = base.getWidth();
                    int h = base.getHeight();

                    BufferedImage combined = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
                    Graphics g = combined.getGraphics();

                    BufferedImage logo = null;

                    switch (c) {
                        case 0:
                            logo = ImageIO.read(new File(sources + "\\" + log + "_BW.png"));
                            break;
                        case 1:
                            logo = ImageIO.read(new File(sources + "\\" + log + "_B.png"));
                            break;
                    }
//            int scaleX = (int) (logo.getWidth() * 0.7);
//            int scaleY = (int) (logo.getHeight() * 0.7);
//            java.awt.Image logo1 = logo.getScaledInstance(scaleX, scaleY, java.awt.Image.SCALE_SMOOTH);

                    String item_mod = itemNo.replace("/2", "").replace("/3", "").replace(".3", "").replace("/4", "").replace(".4", "").replace("/5", "").replace(".5", "");

                    AttributedString word = new AttributedString(item_mod);
                    int item_l = item.length();
                    if (item_l < 7) {
                        word.addAttribute(TextAttribute.FONT, new Font("Calibri", Font.BOLD, 40));
                        word.addAttribute(TextAttribute.FOREGROUND, Color.BLACK);
                    } else {
                        word.addAttribute(TextAttribute.FONT, new Font("Calibri", Font.BOLD, 25));
                        word.addAttribute(TextAttribute.FOREGROUND, Color.BLACK);
                    }

                    AttributedString wattage = new AttributedString(wat);
                    wattage.addAttribute(TextAttribute.FONT, new Font("Calibri", Font.BOLD, 56));
                    wattage.addAttribute(TextAttribute.FOREGROUND, Color.BLACK);
                    int wat_pos_s = 0;
                    int wat_pos_w = 0;
                    switch (wat.length() - 1) {
                        case 1:
                            wat_pos_s = 120;
                            wat_pos_w = 145;
                            break;
                        case 2:
                            wat_pos_s = 85;
                            wat_pos_w = 110;
                            break;
                        case 3:
                            wat_pos_s = 50;
                            wat_pos_w = 75;
                            break;
                        case 4:
                            wat_pos_s = 25;
                            wat_pos_w = 40;
                            break;
                    }

                    //combined.createGraphics().drawImage(combined, 0, 0, Color.YELLOW, null);
                    // paint both images, preserving the alpha channels
                    g.drawImage(base, 0, 0, null);
                    switch (s) {
                        case 0:
                            g.drawString(wattage.getIterator(), wat_pos_s, 720);
                            break;

                        case 1:

                            //g.drawImage(logo,30,170,scaleX,scaleY, null);
                            g.drawImage(logo, 30, 170, null);
                            g.drawString(word.getIterator(), 260, 235);
                            g.drawString(wattage.getIterator(), wat_pos_w, 851);
                            break;
                    }

                    int progress = (int) (50 * s + 25 * c + 25);
                    jProgressBar1.setValue(progress);
                    Rectangle progressRect = jProgressBar1.getBounds();
                    progressRect.x = 0;
                    progressRect.y = 0;
                    jProgressBar1.paintImmediately(progressRect);

                    g.dispose();
//            g2d.dispose();

                    // Save as new image
                    if (!subdir.exists()) {
                        int n = JOptionPane.showConfirmDialog(null, "Would you like to create" + subdir + " folder?", "Folder doesn't exist for this item !!!",
                                JOptionPane.YES_NO_OPTION);
                        if (n == JOptionPane.YES_OPTION) {
                            subdir.mkdir();
                        }
                    }
                    if (subdir.exists()) {
                        ImageIO.write(combined, "PNG", output);
                        if (s == 1 && c == 1) {
                            try {
                                inPut = new FileInputStream(output);
                                outPut = new FileOutputStream(new File(subdir + "\\HR_" + sap + "_9.jpg"));
                                byte[] buf = new byte[1024];
                                int bytesRead;
                                while ((bytesRead = inPut.read(buf)) > 0) {
                                    outPut.write(buf, 0, bytesRead);
                                }
                            } finally {
                                inPut.close();
                                outPut.close();
                            }
                        }
                    }

                    if (PDFCheckBox.isSelected()) {
                        try {
                            String pdf = null;
                            switch (c) {
                                case 0:
                                    pdf = subdir + "\\Energylabel_" + sap + "_" + size[s] + color[c] + ".pdf";
                                    break;
                                case 1:
                                    pdf = subdir + "\\Energylabel_" + sap + "_" + size[s] + "_C.pdf";
                                    break;
                            }

                            Image image = Image.getInstance(output.toString());
                            image.scalePercent((float) 24.01);
                            float image_w = image.getScaledWidth();
                            float image_h = image.getScaledHeight();
                            com.itextpdf.text.Rectangle rect = new com.itextpdf.text.Rectangle(image_w, image_h);
                            Document document = new Document();
                            document.setPageSize(rect);
                            FileOutputStream fos = new FileOutputStream(pdf);
                            PdfWriter writer = PdfWriter.getInstance(document, fos);
                            writer.open();
                            document.open();
                            image.setAbsolutePosition(0, 0);
                            document.add(image);
                            document.close();
                            writer.close();
                        } catch (Exception i1) {
                            i1.printStackTrace();
                        }
                    }
                }
            }
            jProgressBar1.setValue(100);
        } else {
            noitem.add(item);
            //JOptionPane.showMessageDialog(null, "There is no data for item "+item);
        }
    }


    private void jItemStartButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jItemStartButtonActionPerformed
        try {
            String item = jItemTextField.getText().toUpperCase();
            destF = productContent + "\\" + sap;
            getData(item, rownr, itemNo, sap, ean, wat, en_w, log, destF);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(ErP.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ErP.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_jItemStartButtonActionPerformed

    private void jItemTextFieldActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jItemTextFieldActionPerformed
        jItemStartButtonActionPerformed(evt);
    }//GEN-LAST:event_jItemTextFieldActionPerformed

    private void jBrowseButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBrowseButtonActionPerformed

        JFileChooser browse = new JFileChooser(mainfolder);
        browse.setDialogTitle("Select excel file with list");
        browse.setFileSelectionMode(JFileChooser.FILES_ONLY);
        browse.showOpenDialog(null);

        jListTextField.setText(browse.getSelectedFile().getPath());
    }//GEN-LAST:event_jBrowseButtonActionPerformed

    private void jListStartButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jListStartButtonActionPerformed
        try {

            int n = JOptionPane.showConfirmDialog(null, "Save energy labels into PRODUCT CONTENT folders?", "Where to save energy labels", JOptionPane.YES_NO_OPTION);
            if (n == JOptionPane.YES_OPTION) {
                destF = productContent + "\\" + sap;
            } else {
                JFileChooser dest = new JFileChooser(mainfolder);
                dest.setDialogTitle("Select destination folder");
                dest.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                dest.showSaveDialog(null);
                destF = dest.getSelectedFile().toString();
            }

            String path = jListTextField.getText();
            FileInputStream fis1 = null;
            fis1 = new FileInputStream(path);
            XSSFWorkbook wb = new XSSFWorkbook(fis1);
            XSSFSheet sheet = wb.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();

            List<String> noitem = new ArrayList<String>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell = cellIterator.next();
                String item = cell.getStringCellValue();
                getData(item, rownr, itemNo, sap, ean, wat, en_w, log, destF);
//                label(item, destF, noitem, rownr, itemNo, sap, ean, wat, en_w, log);
            }

            if (noitem.size() > 0) {
                JOptionPane.showMessageDialog(null, "Generator didn't find data for folowing items: " + noitem.toString());
            }
            if (!destF.substring(destF.length() - 4).equals("null")) {
                File subdir = new File(destF + "\\");
                Desktop desktop = Desktop.getDesktop();
                desktop.open(subdir);
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ErP.class.getName()).log(Level.SEVERE, null, ex);

        } catch (IOException ex) {
            Logger.getLogger(ErP.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_jListStartButtonActionPerformed

    private void jItemRadioButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jItemRadioButtonActionPerformed
        jItemTextField.requestFocus();
    }//GEN-LAST:event_jItemRadioButtonActionPerformed

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
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ErP.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ErP().setVisible(true);

            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JCheckBox PDFCheckBox;
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.JButton jBrowseButton;
    private javax.swing.JLabel jItemLabel;
    private javax.swing.JRadioButton jItemRadioButton;
    private javax.swing.JButton jItemStartButton;
    private javax.swing.JTextField jItemTextField;
    private javax.swing.JLabel jListLabel;
    private javax.swing.JRadioButton jListRadioButton;
    private javax.swing.JButton jListStartButton;
    private javax.swing.JTextField jListTextField;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JProgressBar jProgressBar1;
    private javax.swing.JLabel jRowCounterLabel;
    private org.jdesktop.beansbinding.BindingGroup bindingGroup;
    // End of variables declaration//GEN-END:variables

}
