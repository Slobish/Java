/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */


package opte;

import java.awt.Component;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.HeadlessException;
import java.io.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


class Filter extends javax.swing.filechooser.FileFilter {
        @Override
        public boolean accept(File file) {
            // Allow only directories, or files with ".txt" extension
            return file.isDirectory() || file.getAbsolutePath().endsWith(".csv");
        }
        @Override
        public String getDescription() {
            // This description will be displayed in the dialog,
            // hard-coded = ugly, should be done via I18N
            return "PTE Files (*.csv)";
        }
    } 

/**
 *
 * @author dell
 */
public class CertificateConverter extends javax.swing.JFrame {
    String SNUM,MODEL,CLASS,VOLTAGE,CURRENT,CONSTANT,DATE,MONTH,YEAR,CERT2300A;
    String MEASUREDATE,REGTYPE,CUSTOMER,PLACE,NAME,WAY,BRAND,CERT2300E;
    String AREA,LOCATION,CTRATIO,VTRATIO,CURPRIM,CURSEC,UNITCONSTANT,TYPE;
    String VOLPRIM,VOLSEC,INOM,UNOM,AVERAGE,DESVIO,TEMP,PREVAUX,PHASE,WIRES;
    String NOTE1,NOTE2,NOTE3,NOTE4,OBSERVATION;
    ArrayList<ArrayList<String>> SAMPLESID = new ArrayList<ArrayList<String>>();
    ArrayList<ArrayList<String>> ERDESV = new ArrayList<ArrayList<String>>();
    List<String> MEASURETIME = new ArrayList<String>();
    GridBagLayout layout = new GridBagLayout();
    FillPanel1 p1;
    FillPanel2 p2;
    FillPanel3 p3;
    int currentPanel=0;
    ResultSet result;
    String USER ="ug000209_opte";   
    String PASS ="DA72toruvu";   
    String DBDIRECTION ="jdbc:mysql://localhost/ug000209_opte";
    /*
     //String USER ="root";
    String USER ="uv7046_eircare";
    //String PASS ="";
    String PASS ="eircare";
    //String DBDIRECTION ="jdbc:mysql://localhost/eircare";
    String DBDIRECTION ="jdbc:mysql://200.68.105.18:3306/eircare";
    */
    FileReader fileReader;
    BufferedReader bufferedReader;
    FileReader intiReader;
    BufferedReader intiBuffered;
    DriverSQL SQL = new DriverSQL(DBDIRECTION,USER,PASS);;
    String selectedPath;
    String selectedOutput;
    String selectedTemplate;
    List<String> plantillas;
 
    public CertificateConverter() 
    {
        initComponents();        
        p1 = new FillPanel1();
        p2 = new FillPanel2(this.SQL);
        p3 = new FillPanel3();
        BackButton.setEnabled(false);
        DynamicPanel.setLayout(layout);
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = 0;
        c.gridy = 0;
        DynamicPanel.add(p1,c);
        c.gridx = 0;
        c.gridy = 0;
        DynamicPanel.add(p2,c);
        c.gridx = 0;
        c.gridy = 0;
        DynamicPanel.add(p3,c);
        p1.setVisible(true);
        p2.setVisible(false);   
        p3.setVisible(false);
       
        try 
        {
            plantillas=textFiles("C:\\Users\\dell\\Documents\\NetBeansProjects\\OPTE\\plantillas");
            for(String file:plantillas)
            {
                Plantillas.addItem(file);
                Plantillas.setSelectedIndex(1);
            }     
        }
        catch (Exception e)
        {
            JOptionPane.showMessageDialog(this, "Error XX02");
        }
       
        
    }

    
    /**
   * return a number of page of document.
   *
   * @param filename
   *             name of the file
   * @return number of pages
   */
  
 
List<String> textFiles(String directory) 
{
    List<String> textFiles = new ArrayList<String>();
    try
    {      
      File dir = new File(directory);
      for (File file : dir.listFiles()) 
      {
        if (file.getName().endsWith((".doc"))) textFiles.add(file.getName());
         
      }
      
    }
    catch(Exception e) 
    {
       JOptionPane.showMessageDialog(this, "Error XX11");   
       
    }
    return textFiles;
    
}
   
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        fileChooser1 = new javax.swing.JFileChooser();
        fileChooser2 = new javax.swing.JFileChooser();
        GlobalPannel = new javax.swing.JPanel();
        SearchFile = new javax.swing.JButton();
        Label1 = new javax.swing.JLabel();
        SearchOutput = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();
        SelectedOutput = new javax.swing.JLabel();
        SelectedFile = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        Plantillas = new javax.swing.JComboBox<>();
        Equipment = new javax.swing.JComboBox<>();
        DBUPLOAD = new javax.swing.JCheckBox();
        ConverterButton = new javax.swing.JButton();
        BackButton = new javax.swing.JButton();
        PanelNumber = new javax.swing.JLabel();
        NextButton = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        ContrastBox = new javax.swing.JCheckBox();
        DynamicPanel = new javax.swing.JPanel();

        fileChooser1.setDialogTitle("Seleccione un archivo PTE");
        fileChooser1.setFileFilter(new Filter());
        fileChooser1.setFileHidingEnabled(true);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        GlobalPannel.setBackground(new java.awt.Color(153, 153, 153));
        GlobalPannel.setToolTipText("");

        SearchFile.setText("Buscar");
        SearchFile.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SearchFileActionPerformed(evt);
            }
        });

        Label1.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        Label1.setText("Tabla");

        SearchOutput.setText("Buscar");
        SearchOutput.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SearchOutputActionPerformed(evt);
            }
        });

        jLabel2.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel2.setText("Salida");

        SelectedOutput.setFont(new java.awt.Font("Tahoma", 0, 10)); // NOI18N
        SelectedOutput.setText("...");

        SelectedFile.setFont(new java.awt.Font("Tahoma", 0, 10)); // NOI18N
        SelectedFile.setText("...");

        jLabel3.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel3.setText("Plantilla");

        Plantillas.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " " }));

        Equipment.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "PTE 2300A", "PTE 2300E", " " }));

        DBUPLOAD.setBackground(new java.awt.Color(153, 153, 153));
        DBUPLOAD.setSelected(true);
        DBUPLOAD.setText("Actualizar base de datos");
        DBUPLOAD.setEnabled(false);

        ConverterButton.setText("Convertir");
        ConverterButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ConverterButtonActionPerformed(evt);
            }
        });

        BackButton.setFont(new java.awt.Font("Tahoma", 1, 8)); // NOI18N
        BackButton.setText("<<");
        BackButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                BackButtonActionPerformed(evt);
            }
        });

        PanelNumber.setText("1/3");

        NextButton.setFont(new java.awt.Font("Tahoma", 1, 8)); // NOI18N
        NextButton.setText(">>");
        NextButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                NextButtonActionPerformed(evt);
            }
        });

        jLabel1.setText("Equipamiento");

        ContrastBox.setSelected(true);
        ContrastBox.setText("Contrastar base de datos");
        ContrastBox.setEnabled(false);
        ContrastBox.setOpaque(false);
        ContrastBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ContrastBoxActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout GlobalPannelLayout = new javax.swing.GroupLayout(GlobalPannel);
        GlobalPannel.setLayout(GlobalPannelLayout);
        GlobalPannelLayout.setHorizontalGroup(
            GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(GlobalPannelLayout.createSequentialGroup()
                .addComponent(BackButton, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(36, 36, 36)
                .addComponent(PanelNumber)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(NextButton))
            .addGroup(GlobalPannelLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(GlobalPannelLayout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addGap(18, 18, 18)
                        .addComponent(SearchOutput)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(SelectedOutput))
                    .addGroup(GlobalPannelLayout.createSequentialGroup()
                        .addComponent(Label1)
                        .addGap(21, 21, 21)
                        .addComponent(SearchFile)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(SelectedFile))
                    .addGroup(GlobalPannelLayout.createSequentialGroup()
                        .addComponent(jLabel3)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(Plantillas, javax.swing.GroupLayout.PREFERRED_SIZE, 97, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(GlobalPannelLayout.createSequentialGroup()
                .addGap(18, 18, 18)
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(ContrastBox)
                    .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addComponent(DBUPLOAD, javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(GlobalPannelLayout.createSequentialGroup()
                            .addComponent(ConverterButton)
                            .addGap(27, 27, 27)))
                    .addGroup(GlobalPannelLayout.createSequentialGroup()
                        .addComponent(jLabel1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(Equipment, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(0, 22, Short.MAX_VALUE))
        );
        GlobalPannelLayout.setVerticalGroup(
            GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(GlobalPannelLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(Label1)
                    .addComponent(SelectedFile)
                    .addComponent(SearchFile))
                .addGap(26, 26, 26)
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(SearchOutput)
                    .addComponent(SelectedOutput))
                .addGap(27, 27, 27)
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(Plantillas, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(30, 30, 30)
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(Equipment, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 18, Short.MAX_VALUE)
                .addComponent(ContrastBox)
                .addGap(18, 18, 18)
                .addComponent(DBUPLOAD)
                .addGap(18, 18, 18)
                .addComponent(ConverterButton)
                .addGap(18, 18, 18)
                .addGroup(GlobalPannelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(BackButton, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(PanelNumber)
                    .addComponent(NextButton, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)))
        );

        javax.swing.GroupLayout DynamicPanelLayout = new javax.swing.GroupLayout(DynamicPanel);
        DynamicPanel.setLayout(DynamicPanelLayout);
        DynamicPanelLayout.setHorizontalGroup(
            DynamicPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 550, Short.MAX_VALUE)
        );
        DynamicPanelLayout.setVerticalGroup(
            DynamicPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(GlobalPannel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(DynamicPanel, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(GlobalPannel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addComponent(DynamicPanel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void SearchFileActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SearchFileActionPerformed

        String aux;
        int returnVal = fileChooser1.showOpenDialog(this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
        File file = fileChooser1.getSelectedFile();
        try 
        {
          selectedPath=file.getAbsolutePath();
          aux=file.getName();
          if(aux.length()>10) SelectedFile.setText(aux.substring(0,10)+".."); 
          else SelectedFile.setText(aux);
        } 
        catch (/*IOException ex*/Exception e)
        {
          JOptionPane.showMessageDialog(this, "ERROR XX03"+file.getAbsolutePath());  
          
        }
          cleanVariables();
          analizeTable();
          fillFields();         
         
        
        
    } else 
    {
        System.out.println("File access cancelled by user.");   
    } 
    /*Select = new JFileChooserDemo();
    Select.setVisible(true);
       */
    }//GEN-LAST:event_SearchFileActionPerformed

    private void ConverterButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ConverterButtonActionPerformed
    if(p2.seen==true && p3.seen==true)
    {
        try
        {
            //DBupload();
            writeDocument();            
        }
        catch(Exception e)
        {
            
        }
        
    }
    else JOptionPane.showMessageDialog(this, "No se han verificado todos los campos.");
    
    }//GEN-LAST:event_ConverterButtonActionPerformed

public void cleanVariables()
{
    try
    {
    p3.cleanTable();
    SNUM=MODEL=CLASS=VOLTAGE=CURRENT=CONSTANT=DATE=MONTH=YEAR=CERT2300A="";
    MEASUREDATE=REGTYPE=CUSTOMER=PLACE=NAME=WAY=BRAND=CERT2300E="";
    AREA=LOCATION=CTRATIO=VTRATIO=CURPRIM=CURSEC=UNITCONSTANT=TYPE="";
    VOLPRIM=VOLSEC=INOM=UNOM=AVERAGE=DESVIO=TEMP=PREVAUX=PHASE=WIRES="";
    NOTE1=NOTE2=NOTE3=NOTE4=OBSERVATION="";
    SAMPLESID.clear();
    ERDESV.clear();
    MEASURETIME.clear();
    }
    catch (Exception e)
    {
        JOptionPane.showMessageDialog(this, "Error XX01");
    }
    
}

public void DBupload()
{
    /*
        if(DBUPLOAD.isSelected())
    {
        try
        {
           if (this.SQL.crearConexion())
                {
                    String query ="";
                    
                    result=this.SQL.ejecutarSQLSelect(query); 
                
                    if(result.next())
                    {
                        
                    }
                     
                }  
           
        }
        catch(Exception e)
        {
          JOptionPane.showMessageDialog(this, "ERROR XX12");
          System.out.println(e);
          
        }
    }  
    */
    
}

public boolean getContrastState()
{
    return ContrastBox.isSelected();
}
public void analizeTable()
{             
        String line; 
        String[] splitLine;
        String aux;
        int state1 = 0;
        int state2 = 0;
        int counter = 0;
        int count = 0;
        int aux2=-1,aux3=0,aux4=-1,aux5=0;
        PREVAUX="";
        
    try
    {   
        fileReader = new FileReader(selectedPath);
        bufferedReader = new BufferedReader(fileReader);
    }
     catch(Exception e)
    {
        JOptionPane.showMessageDialog(this, "ERROR XX05:"+selectedPath);
    } 
    try
    {   
        while((line = bufferedReader.readLine()) != null) 
        {       
            
                if(! line.equals("")) 
                {
                    splitLine=line.split(";");
                    if(splitLine.length>1)
                    {
                        switch (state1)
                        {
                            case 0:
                                if(splitLine[1].equals("LP_ID")) state1=1;  
                                break;
                            case 1:
                                if (splitLine[1].equals("DATETIME")&& splitLine.length>2)
                                {

                                    MEASUREDATE=(splitLine[2].split(" "))[0];
                                    DATE=MEASUREDATE.split("/")[0];
                                    MONTH=MEASUREDATE.split("/")[1];
                                    YEAR=MEASUREDATE.split("/")[2];
                                    MEASURETIME.add((splitLine[2].split(" "))[1]);                               
                                }
                                else if(splitLine[1].equals("METERERROR") && splitLine.length>2)
                                {
                                    state1=2;
                                    ERDESV.add(++aux4,new ArrayList<String>());
                                    ERDESV.get(aux4).add(aux5++,splitLine[2]);

                                }
                                else state1=0;
                                break;
                            case 2:
                                if(splitLine[1].equals("METERSTDEV") && splitLine.length>2)
                                {
                                    ERDESV.get(aux4).add(aux5,splitLine[2]);
                                    aux5=0;
                                }
                                state1=0;
                                break;
                        }                

                        switch(state2)
                        {
                            case 0:
                                if(splitLine[1].equals("SAMP_LP_ID"))
                                {
                                    if(!(splitLine[2].equals(PREVAUX)))
                                    {
                                       PREVAUX=splitLine[2];
                                       SAMPLESID.add(++aux2, new ArrayList<String>());                                   
                                       aux3=0;                                   
                                    } 
                                    state2=1;                              
                                }
                                break;
                            case 1:
                                if(splitLine[1].equals("SAMP_ERROR"))
                                {
                                    SAMPLESID.get(aux2).add(aux3++,splitLine[2]);                                
                                    state2=0;
                                }
                                break;
                        }
                        if(splitLine[1].equals("SNUM") && splitLine.length>2)       SNUM=splitLine[2];
                        if(splitLine[1].equals("TYPE") && splitLine.length>2)       MODEL=splitLine[2];
                        if(splitLine[1].equals("CLASS")&& splitLine.length>2)       
                        {
                           CLASS=splitLine[2];
                           if(CLASS.length()>3)CLASS=CLASS+CLASS.substring(0,3);
                        }
                        if(splitLine[1].equals("VOLTAGE")&& splitLine.length>2)     VOLTAGE=splitLine[2];
                        if(splitLine[1].equals("CURRENT")&& splitLine.length>2)     CURRENT=splitLine[2];
                        if(splitLine[1].equals("CONSTANT")&& splitLine.length>2)    CONSTANT=splitLine[2];
                        if(splitLine[1].equals("CUSTOMER")&& splitLine.length>2)    CUSTOMER=splitLine[2];
                        if(splitLine[1].equals("ADDRESS")&& splitLine.length>2)     PLACE=splitLine[2];
                        if(splitLine[1].equals("NOTE1")&& splitLine.length>2)       NAME=splitLine[2];
                        if(splitLine[1].equals("NOTE2")&& splitLine.length>2)       WAY=splitLine[2];
                        if(splitLine[1].equals("NOTE3")&& splitLine.length>2)       BRAND=splitLine[2];
                        if(splitLine[1].equals("AREA")&& splitLine.length>2)        AREA=splitLine[2];
                        if(splitLine[1].equals("LOCATION")&& splitLine.length>2)    LOCATION=splitLine[2];
                        if(splitLine[1].equals("CTRATIO")&& splitLine.length>2)     CTRATIO=splitLine[2];
                        if(splitLine[1].equals("VTRATIO")&& splitLine.length>2)     VTRATIO=splitLine[2];
                        if(splitLine[1].equals("CURPRIM")&& splitLine.length>2)     CURPRIM=splitLine[2];
                        if(splitLine[1].equals("CURSEC")&& splitLine.length>2)      CURSEC=splitLine[2];
                        if(splitLine[1].equals("VOLPRIM")&& splitLine.length>2)     VOLPRIM=splitLine[2];
                        if(splitLine[1].equals("VOLSEC")&& splitLine.length>2)      VOLSEC=splitLine[2];
                        if(splitLine[1].equals("INOM")&& splitLine.length>2)        INOM=splitLine[2];
                        if(splitLine[1].equals("UNOM")&& splitLine.length>2)        UNOM=splitLine[2];
                        if(splitLine[1].equals("METERERROR")&& splitLine.length>2)  AVERAGE=splitLine[2];
                        if(splitLine[1].equals("METERSTDEV")&& splitLine.length>2)  DESVIO=splitLine[2];          
                        if(splitLine[1].equals("CONSTUNIT")&& splitLine.length>2)   UNITCONSTANT=splitLine[2];
                        if(splitLine[1].equals("CONTYPE")&& splitLine.length>2)
                        {
                            
                            PHASE=splitLine[2].substring(0,1);
                            WIRES=splitLine[2].substring(2,3);
                           
                        }
                        
                    }
                }
            }
        }
        catch (Exception e)
            {
                JOptionPane.showMessageDialog(this, "ERROR XX06");
            }
    try
    {
        bufferedReader.close();
    }
    catch(Exception e)
    {
        JOptionPane.showMessageDialog(this, "ERROR XX010");
    }
        }       

public void writeDocument()
{      HWPFDocument doc = null;
       FileInputStream fis=null;
       POIFSFileSystem fs=null;
    try
        {  
            TYPE=p1.getType();
            SNUM=p1.getSNumber();
            CURRENT=p1.getCurrent();
            VOLTAGE=p1.getVoltage();
            CLASS=p1.getClassBox();
            PHASE=p1.getPhase();
            WIRES=p1.getWire();
            CONSTANT=p1.getConstant();
            UNITCONSTANT=p1.getConsUnit();
            BRAND=p1.getBrand();
            MODEL=p1.getModel();
            NAME=p1.getNameCAM();
            CUSTOMER=p2.getCustomer();
            PLACE=p2.getPlace();
            AREA=p2.getArea();
            LOCATION=p2.getLocal();
            VTRATIO=p1.getVTRATIO();
            CTRATIO=p1.getCTRATIO();
            TEMP=p1.getTemp();
            INOM=p3.getINOM();
            UNOM=p3.getUNOM();
            WAY=p1.getWay();
            NOTE1=p3.getNote1();
            NOTE2=p3.getNote2();
            NOTE3=p3.getNote3();
            NOTE4=p3.getNote4();   
        
    try
    {
        selectedTemplate="C:\\Users\\dell\\Documents\\NetBeansProjects\\OPTE\\plantillas\\"+Plantillas.getItemAt(Plantillas.getSelectedIndex());
        fis = new FileInputStream(selectedTemplate );
        fs = new POIFSFileSystem(fis);
        doc = new HWPFDocument(fs);
    }
    catch(Exception e)
    {
        JOptionPane.showMessageDialog(this, "ERROR XX07");
    }        
    try
    {
            Range range = doc.getRange();
            range.sanityCheck();            
            range.replaceText("%SNUM%",SNUM);
            range.replaceText("%MEASUREDATE%",MEASUREDATE);
            range.replaceText("%MONTH%",MONTH);
            range.replaceText("%YEAR%",YEAR);
            range.replaceText("%PHASE%",PHASE);
            range.replaceText("%WIRES%",WIRES);
            range.replaceText("%CLASS%",CLASS);
            range.replaceText("%METERTYPE%",TYPE);
            range.replaceText("%CONSTANT%",CONSTANT);
            range.replaceText("%UNITCONSTANT%",UNITCONSTANT);
            range.replaceText("%CURRENT%",CURRENT);
            range.replaceText("%VOLTAGE%",VOLTAGE);
            range.replaceText("%BRAND%",BRAND);
            range.replaceText("%MODEL%",MODEL);
            range.replaceText("%NAME%",NAME);
            range.replaceText("%CUSTOMER%",CUSTOMER);
            range.replaceText("%ADDRESS%",PLACE);
            range.replaceText("%AREA%",AREA);
            range.replaceText("%LOCATION%",LOCATION);
            range.replaceText("%MODELPTE%",getInstrument());
            range.replaceText("%CLASSPTE%",getInstrumentClass());
            range.replaceText("%SNUMPTE%",getInstrumentSNUM());
            range.replaceText("%CERTIFICATE%",getInstrumentCertificate());
            range.replaceText("%VTRATIO%",VTRATIO);
            range.replaceText("%CTRATIO%",CTRATIO);
            range.replaceText("%MEAINIT%",MEASURETIME.get(0));
            range.replaceText("%MEAFINISH%",MEASURETIME.get(MEASURETIME.size()-1));
            range.replaceText("%TEMP%",TEMP);
            range.replaceText("%COMPERDIDAS%",p1.getCompensation());
            range.replaceText("%INOM%",INOM);
            range.replaceText("%UNOM%",UNOM);
            range.replaceText("%INOM10%",calcPercentage(INOM,(float)0.1));
            range.replaceText("%INOM5%",calcPercentage(INOM,(float)5/100));
            range.replaceText("%INOM1%",calcPercentage(INOM,(float)1/100));
            range.replaceText("%WAY%",WAY);
            range.replaceText("%VOLPRIM%",p1.getVTRATIO().substring(0,p1.getVTRATIO().indexOf("/")+1));
            range.replaceText("%VOLSEC%",p1.getVTRATIO().substring(p1.getVTRATIO().indexOf("/")+1));
            range.replaceText("%CURPRIM%",p1.getCTRATIO().substring(0,p1.getCTRATIO().indexOf("/")+1));
            range.replaceText("%CURSEC%",p1.getCTRATIO().substring(0,p1.getCTRATIO().indexOf("/")+1));
            range.replaceText("%OBSERVATION%",OBSERVATION);
            range.replaceText("%CUSTOMINFORMATION1%",NOTE1);
            range.replaceText("%CUSTOMINFORMATION2%",NOTE2);
            range.replaceText("%CUSTOMINFORMATION3%",NOTE3);
            range.replaceText("%CUSTOMINFORMATION4%",NOTE4);
            
            
            for(int help=1;help<=p3.getMeasurements();help++)
            {
               for(int help2=1;help2<=p3.getSamples();help2++)
               {
                   String aux1= ""+help;
                   String aux2= ""+help2;
                   range.replaceText("%SAMPLE"+aux1+aux2+"%",SAMPLESID.get(help-1).get(help2-1).substring(0,7));
               }
            }
            for(int help=1;help<=p3.getMeasurements();help++)
            {
               String aux1= ""+help;
               range.replaceText("%RESULT"+aux1+"1%",ERDESV.get(help-1).get(0).substring(0,7));
               range.replaceText("%RESULT"+aux1+"2%",ERDESV.get(help-1).get(1).substring(0,7));
            }
    }
    catch (Exception e)
    {
        JOptionPane.showMessageDialog(this, "ERROR XX08");
    }
          try
          {             
            FileOutputStream fos = new FileOutputStream(selectedOutput);
            doc.write(fos);
            fis.close();
            fos.close();
          }
          catch(Exception e)
          {
              JOptionPane.showMessageDialog(this, "ERROR XX09");
          }
        }
   
    catch(Exception e)
    {
        JOptionPane.showMessageDialog(this, "Error XX04");
    }
}
    private void SearchOutputActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SearchOutputActionPerformed
        String aux;
        int returnVal = fileChooser2.showOpenDialog(this);
        if (returnVal == JFileChooser.APPROVE_OPTION) 
        {
        File file = fileChooser2.getSelectedFile();
        try {
         
          selectedOutput=file.getAbsolutePath();
          selectedOutput=selectedOutput+".doc";
          aux=file.getName();
          if(aux.length()>10) SelectedOutput.setText(aux.substring(0,10)+".."); 
          else SelectedOutput.setText(aux);          
          
        } 
        catch (Exception e)
        {
          JOptionPane.showMessageDialog(this, "No se pudo abrir el archivo "+file.getAbsolutePath());         
        }
        } 
        else 
        {
            System.out.println("File access cancelled by user.");   
        } 
    }//GEN-LAST:event_SearchOutputActionPerformed

    public String getInstrumentClass()
    {
        if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300A")) return "0.5";
                else if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300E")) return "0.2";
        return "";
    }
    public String getInstrument()
    {
        if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300A")) return "2300A";
                else if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300E")) return "2300E";
        return "";
    }
    public String getInstrumentSNUM()
    {
        if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300A")) return "2613100101";
                else if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300E")) return "2615040136";
        return "";
    }
    public String getInstrumentCertificate()
    {   
        if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300A")) return "INTI N° FM-102-18194";
                else if(Equipment.getItemAt(Equipment.getSelectedIndex()).equals("PTE 2300E")) return "NO";
        return "";
        
    }
    private void BackButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_BackButtonActionPerformed
       switch(currentPanel)
       {
           case 1:
               p1.setVisible(true);
               p2.setVisible(false);
               PanelNumber.setText("1/3");
               BackButton.setEnabled(false);
               NextButton.setEnabled(true);
               currentPanel=0;
               break;
           case  2:
               p3.setVisible(false);
               p2.setVisible(true);
               NextButton.setEnabled(true);
               PanelNumber.setText("2/3");
               currentPanel=1;
               break;
       }
    }//GEN-LAST:event_BackButtonActionPerformed
public String calcPercentage(String str,float percentage)
{
   
     Float number = new Float (str);
     float realnumber= number.floatValue();
     realnumber=realnumber*percentage;
     number= new Float(realnumber);
     return number.toString();
}
    private void NextButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_NextButtonActionPerformed
      
        switch(currentPanel)
       {
           case 0:
               p1.setVisible(false);
               p2.setVisible(true);
               p2.seen=true;
               PanelNumber.setText("2/3");
               BackButton.setEnabled(true);
               currentPanel=1;
               break;
           case 1:
               p3.setVisible(true);
               p2.setVisible(false);
               p3.seen=true;
               PanelNumber.setText("3/3");
               NextButton.setEnabled(false);
               currentPanel=2;
               break;
               
       }
    }//GEN-LAST:event_NextButtonActionPerformed

   
    private void ContrastBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ContrastBoxActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_ContrastBoxActionPerformed



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
            java.util.logging.Logger.getLogger(CertificateConverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(CertificateConverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(CertificateConverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(CertificateConverter.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new CertificateConverter().setVisible(true);
        });
    }
    
   public void fillFields() 
   {
       /*
       if (ContrastBox.isSelected()) 
        {
            
            String query= "SELECT * FROM `Medidores` WHERE `Numero de serie` LIKE "+SNUM;       
            try
            {        
                if (this.SQL.crearConexion())
                {
                     print("Conexion exitosa");
                }    
                
                result=this.SQL.ejecutarSQLSelect(query);
                
                if(result.first())
                    {
                        if(confirmPopUp("¿Desea completar los campos con la información en la base de datos?") == 0)
                        {
                            databaseContrast();              
                        }
                        else
                        {
                            boolean test=true;
                            p1.setSNumber(SNUM);
                            p1.setBrand(BRAND);
                            p1.setConstant(CONSTANT);
                            p1.setModel(MODEL);
                            p1.setNameCAM(NAME);
                            p1.setConsUnit(UNITCONSTANT);
                            p1.setClassBox(CLASS);
                            p1.setPhase(PHASE);
                            p1.setWire(WIRES);
                            p3.setINOM(INOM);
                            p3.setUNOM(UNOM);
                            p3.setStartTime(MEASURETIME.get(0));
                            p3.setFinishTime(MEASURETIME.get(MEASURETIME.size()-1));
                            for(int i=0;i<p3.getMeasurements();i++)
                            {

                                Float numero = new Float(ERDESV.get(i).get(0));          
                                float numero1= numero.floatValue();         
                                numero=new Float(CLASS);
                                if(numero1>= numero.floatValue())
                                {
                                    test=false;
                                }
                            }
                            p3.fillTable(ERDESV);
                            p3.setResult(test);
                        }
           
                    }
                 else
                {
                    JOptionPane.showMessageDialog(this,"No se encontró registro en la base de datos asociado al numero de serie: "+ SNUM);
           
                }
                 
            }
             catch(Exception e)
            {
                System.out.print(e);
                JOptionPane.showMessageDialog(this, "ERROR XX12");  
            }           
        }
       else
       {
           */
       try
       {
       boolean test=true;
       p1.setSNumber(SNUM);
       System.out.println("1");
       p1.setBrand(BRAND);
       System.out.println("1");
       p1.setConstant(CONSTANT);
       System.out.println("1");
       p1.setModel(MODEL);
       System.out.println("1");
       p1.setNameCAM(NAME);
       System.out.println("1");
       p1.setConsUnit(UNITCONSTANT);
       System.out.println("2");
      // p1.setClassBox(CLASS);
       System.out.println("2");
       p1.setPhase(PHASE);
       System.out.println("1");
       p1.setWire(WIRES);
       System.out.println("1");
       p3.setINOM(INOM);
       System.out.println("1");
       p3.setUNOM(UNOM);
       System.out.println("1");
       p3.setStartTime(MEASURETIME.get(0));
       p3.setFinishTime(MEASURETIME.get(MEASURETIME.size()-1));
       for(int i=0;i<p3.getMeasurements();i++)
       {
           
           Float numero = new Float(ERDESV.get(i).get(0));          
           float numero1= numero.floatValue();         
           numero=new Float(CLASS);
           if(numero1>= numero.floatValue())
           {
               test=false;
           }
       }
       p3.fillTable(ERDESV);
       p3.setResult(test);
       } 
       catch(Exception e)
        {
           System.out.println(e);
           JOptionPane.showMessageDialog(this, "ERROR XX14");  
        }
       }
          
   public void databaseContrast() 
   {
       
        String query = "SELECT * FROM medidores INNER JOIN areas ON medidores.ID_Area = areas.ID_Area INNER JOIN"
                + " centrales ON centrales.ID_Central=areas.ID_Central INNER JOIN empresas ON "
                + "centrales.ID_Empresa=empresas.ID_Empresa WHERE medidores.`Numero de serie` "
                + "LIKE '"+SNUM+"'";
        try
        {
               result=this.SQL.ejecutarSQLSelect(query); 
               if(result.next())
               {
                    p1.setSNumber(result.getString("Numero de Serie"));
                    p1.setNameCAM(result.getString("Nombre_Camesa"));
                    p1.setBrand(result.getString("Marca"));
                    p1.setModel(result.getString("Modelo"));
                    p1.setConstant(result.getString("Constante"));
                    p2.setCompany(result.getString("Nombre_Empresa"));
                    p2.setLocation(result.getString("Localidad_Central"));
                    p2.setPlace(result.getString("Nombre_Central"));
                    p2.setEmail(result.getString("Email"));
                    p2.setPhone(result.getString("Telefono"));
                    p1.setVoltage(result.getString("Voltaje"));
                    p1.setCurrent(result.getString("Corriente"));
                    p1.setPhase(result.getString("Fases"));
                    p1.setWire(result.getString("Hilos"));
                    p1.setClassBox(result.getString("Clase").substring(0,result.getString("Clase").indexOf("S")));

                 }
        }
        catch(Exception e)
        {
            JOptionPane.showMessageDialog(this, "ERROR XX13");  
            print(e.toString());
        }
      
   }

   public static int confirmPopUp(String theMessage) 
   {
    int result = JOptionPane.showConfirmDialog((Component) null, theMessage,
        "alert", JOptionPane.OK_CANCEL_OPTION);
    return result;
   }
   public void print(String text)
   {
       System.out.println(text);
   }
  

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton BackButton;
    private javax.swing.JCheckBox ContrastBox;
    private javax.swing.JButton ConverterButton;
    private javax.swing.JCheckBox DBUPLOAD;
    private javax.swing.JPanel DynamicPanel;
    private javax.swing.JComboBox<String> Equipment;
    private javax.swing.JPanel GlobalPannel;
    private javax.swing.JLabel Label1;
    private javax.swing.JButton NextButton;
    private javax.swing.JLabel PanelNumber;
    private javax.swing.JComboBox<String> Plantillas;
    private javax.swing.JButton SearchFile;
    private javax.swing.JButton SearchOutput;
    private javax.swing.JLabel SelectedFile;
    private javax.swing.JLabel SelectedOutput;
    private javax.swing.JFileChooser fileChooser1;
    private javax.swing.JFileChooser fileChooser2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    // End of variables declaration//GEN-END:variables
}
