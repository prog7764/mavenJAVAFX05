package com.steph.mavenjavafx05;

import static com.steph.mavenjavafx05.sSQLiteManager.intToLetter;
import static com.steph.mavenjavafx05.sSQLiteManager.letterToInt;
import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import static java.util.Collections.list;
import java.util.HashSet;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Set;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.Separator;
import javafx.scene.control.Tab;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import javax.swing.JFileChooser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.sql.*;
import java.util.Arrays;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Workbook;

public class FXMLController implements Initializable {
    
    List<ChoiceBox<String>> m_ChoiceBoxSalles = new ArrayList<>();
    private String  m_Fichier_Etudiants = "";
    private String  m_Fichier_TiersTemps = "";
    private String  m_Fichier_Salles = "";
    
    private String  m_DS1;
    private String  m_Annee_DS1;
    private String  m_DS2;
    private String  m_Annee_DS2;
    
    private int     m_DateSeed;

    private String  m_Fichier_EDT_3IF;
    private String  m_Fichier_EDT_4IF;
    private String  m_Fichier_EDT_5IF;
    
    private String  m_FichierInfoComplementaire;
    private String  m_Fichier_Maquette;
 
    @FXML
    private Label label;
    
    
    @FXML
    private Tab FXID_TAB_1_DS;
    
    @FXML
    private Tab FXID_TAB_2_HEURES;
    @FXML
    private Tab FXID_TAB_3_NOTES;
    @FXML
    private Tab FXID_TAB_9_PARAMETRES;
    @FXML
    private TextField FXID_TXF_9_ETUDIANTS;
    @FXML
    private TextField FXID_TXF_9_TIERSTEMPS;
    @FXML
    private Button FXID_BTN_9_ETUDIANTS;
    @FXML
    private Button FXID_BTN_9_TIERSTEMPS;
    @FXML
    private TextField FXID_TXF_9_SALLES;
    @FXML
    private Button FXID_BTN_9_SALLES;
    private TextField FXID_TXF_9_RESERVATIONSALLES;
    @FXML
    private DatePicker FXID_DPR_1_DATE;
    @FXML
    private Button FXID_BTN_1_INIT;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DS1; 
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DS2;
    @FXML
    private TextArea FXID_TXA_1_MESSAGES;
    @FXML
    private TextField FXID_TXF_1_RESULTAT;
    @FXML
    private Button FXID_BTN_1_RESULTAT;
    @FXML
    private Button FXID_BTN_1_GO;
    @FXML
    private Button FXID_BTN_QUIT;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE00;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE01;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE02;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE06;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE07;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE03;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE08;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE04;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE09;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_SALLE05;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DELTATT_DS1;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DEBUT_DS2;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DEBUT_DS1;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DELTATT_DS2;
    @FXML
    private TextArea FXID_TXA_1_MESSAGES_DS1;
    @FXML
    private TextArea FXID_TXA_1_MESSAGES_DS2;
    @FXML
    private Button FXID_BTN_1_GOTEST;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DUREE_DS1;
    @FXML
    private ChoiceBox<String> FXID_CCB_1_DUREE_DS2;
    private String m_Fichier_Etudiants_Echanges;
    @FXML
    private TextField FXID_TXF_9_ETUDIANTS_ECHANGES;
    @FXML
    private Button FXID_BTN_9_ETUDIANTS_ECHANGES;
    @FXML
    private Button FXID_BTN_2_INIT;
    @FXML
    private TextArea FXID_TXA_2_MESSAGES;
    @FXML
    private TextField FXID_TXF_9_CRENEAUX_3IF;
    @FXML
    private Button FXID_BTN_9_CRENEAUX_3IF;
    @FXML
    private TextField FXID_TXF_9_CRENEAUX_4IF;
    @FXML
    private Button FXID_BTN_9_CRENEAUX_4IF;
    @FXML
    private TextField FXID_TXF_9_CRENEAUX_5IF;
    @FXML
    private Button FXID_BTN_9_CRENEAUX_5IF;
    @FXML
    private Button FXID_BTN_2_CALCUL_HEURES;
    @FXML
    private TextField FXID_TXF_3_XLS_SCOL;
    @FXML
    private Button FXID_BTN_3_XLS_SCOL;
    @FXML
    private Button FXID_BTN_3_GO;
    @FXML
    private TextArea FXID_TXA_3_MESSAGES;
    
    
    private void addMessage(int target, String message) {
        addMessage(target, message, false);
    }
    private void addMessage(int target, String message, boolean debug) {
        
        switch (target) {
            case 1: 
                FXID_TXA_1_MESSAGES.appendText(message);
                break;
            case 2:
                FXID_TXA_2_MESSAGES.appendText(message);
                break;
            case 3:
                FXID_TXA_3_MESSAGES.appendText(message);
                break;
            default :
        }
    }
    
    
    
    private void handleButtonAction(ActionEvent event) {
        System.out.println("You clicked me!");
        label.setText("Hello World!");
    }
    
    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // TODO
        //m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE00);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE01);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE02);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE03);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE04);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE05);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE06);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE07);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE08);
        m_ChoiceBoxSalles.add(FXID_CCB_1_SALLE09);
        
//        try {
//            // Extraction des paramêtres par défaut
//            sExcelFileManager sEFM = new sExcelFileManager("./settings.xlsx");
//            m_Fichier_Etudiants = sEFM.getDataFromCurrentSheetFirstColumnValue("Fichier_Etudiants", "Valeur");
//            m_Fichier_TiersTemps = sEFM.getDataFromCurrentSheetFirstColumnValue("Fichier_TiersTemps", "Valeur");
//            m_Fichier_Salles = sEFM.getDataFromCurrentSheetFirstColumnValue("Fichier_Salles", "Valeur");
//            
//        } catch (IOException ex) {
//            Logger.getLogger(FXMLController.class.getName()).log(Level.SEVERE, null, ex);
//        } catch (InvalidFormatException ex) {
//            Logger.getLogger(FXMLController.class.getName()).log(Level.SEVERE, null, ex);
//        } catch (sExcelFileManagerException ex) {
//            Logger.getLogger(FXMLController.class.getName()).log(Level.SEVERE, null, ex);
//        }
        m_Fichier_Etudiants="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/Etudiants_IF_2017-2018-IP.xlsx";
        m_Fichier_Etudiants_Echanges="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/Etudiants_Echanges_2017-2018-IP.xlsx";
        m_Fichier_TiersTemps="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/IF-TiersTemps-2017-2018.xlsx";
        m_Fichier_Salles="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/Salles-BlaisePascal.xls";
        FXID_TXF_9_ETUDIANTS.setText(m_Fichier_Etudiants);
        FXID_TXF_9_ETUDIANTS_ECHANGES.setText(m_Fichier_Etudiants_Echanges);
        FXID_TXF_9_TIERSTEMPS.setText(m_Fichier_TiersTemps);
        FXID_TXF_9_SALLES.setText(m_Fichier_Salles);
        
        m_Fichier_EDT_3IF ="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/3IF_EdT_2017-2018.xlsx";
        m_Fichier_EDT_4IF ="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/4IF_EdT_2017-2018.xlsx";
        m_Fichier_EDT_5IF ="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/5IF_EdT_2017-2018.xlsx";
        FXID_TXF_9_CRENEAUX_3IF.setText(m_Fichier_EDT_3IF);
        FXID_TXF_9_CRENEAUX_4IF.setText(m_Fichier_EDT_4IF);
        FXID_TXF_9_CRENEAUX_5IF.setText(m_Fichier_EDT_5IF);
    
        m_FichierInfoComplementaire="/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/IF_Complements_2017-2018.xlsx";
        m_Fichier_Maquette = "/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/Maquette-3IF-4IF-5IF-version 2017-2018.xlsx";

    }    

    
    
    
    
    
    
    
    
    
    
    @FXML
    private void OnAction_BTN_9_ETUDIANTS(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            FXID_TXF_9_ETUDIANTS.setText(selectedFile.getPath());
        }
    }

    @FXML
    private void OnAction_BTN_9_TIERSTEMPS(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            FXID_TXF_9_TIERSTEMPS.setText(selectedFile.getPath());
        }
    }

    @FXML
    private void OnAction_BTN_9_SALLES(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            FXID_TXF_9_SALLES.setText(selectedFile.getPath());
        }
    }

    private void OnAction_BTN_9_RESERVATIONSALLES(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            FXID_TXF_9_RESERVATIONSALLES.setText(selectedFile.getPath());
        }
    }

    
    
    
    
    
    
    
    
    
    
    
    
    @FXML
    private void OnAction_DPR_1_DATE(ActionEvent event) {
    }

    @FXML
    private void OnAction_BTN_1_INIT(ActionEvent event) throws IOException, InvalidFormatException, sExcelFileManagerException {
        addMessage(1,"----- INIT ------------\n");
        boolean bToutBon=true;
        
//        // Création de base de données SQLite
//        sEffectIFSGBD effectIF = new sEffectIFSGBD();
//        if (effectIF.aUnProbleme()) {
//            addMessage(1,("PROBLEME dans effectIF:\n" + effectIF.getMessage());
//            bToutBon=false;
//        }


        // Vérification des paramètres saisis
        
        // Liste des EC
        addMessage(1,"Mise à jour de la liste des EC : ");
        if ("".equals(FXID_TXF_9_ETUDIANTS.getText())){
            addMessage(1,"PROBLEME \n-> Fichier des étudiants non précisé -> 9. Parametres\n");
            bToutBon=false;
        }
        else {
            sExcelFileManager sEFM = new sExcelFileManager(FXID_TXF_9_ETUDIANTS.getText());
            List<String> listTemp = sEFM.getColumnFromHeader("MEC_CODE");
            listTemp.add("- aucun -");
            // Suppression des doublons
            Set<String> set = new HashSet<String>();
            set.addAll(listTemp);
            List<String> listEC = new ArrayList<String>(set);
            Collections.sort(listEC);
            FXID_CCB_1_DS1.setItems(FXCollections.observableArrayList( listEC));
            FXID_CCB_1_DS1.getSelectionModel().select(0);
            FXID_CCB_1_DS2.setItems(FXCollections.observableArrayList( listEC));
            FXID_CCB_1_DS2.getSelectionModel().select(0);
            
            addMessage(1,"OK\n");
        }
        
        // Liste des SALLES
        addMessage(1,"Mise à jour de la liste des Salles : ");
        if ("".equals(FXID_TXF_9_SALLES.getText())){
            addMessage(1,"PROBLEME \n-> Fichier des salles non précisé -> 9. Parametres\n");
            bToutBon=false;
        }
        else {
            sExcelFileManager sEFM = new sExcelFileManager(FXID_TXF_9_SALLES.getText());
            List<String> listSalles = sEFM.getColumnFromHeader("Nom");
            for (ChoiceBox<String> cb : m_ChoiceBoxSalles) {
                cb.setItems(FXCollections.observableArrayList( listSalles));
                cb.getSelectionModel().select(0);
            } 
            FXID_CCB_1_SALLE00.setItems(FXCollections.observableArrayList( listSalles));
            FXID_CCB_1_SALLE00.getSelectionModel().select(0);

            addMessage(1,"OK\n");
        }
        
        // Liste des ETUDIANTS TIERS TEMPS
        addMessage(1,"Mise à jour de la liste des Etudiants Tiers Temps : ");
        if ("".equals(FXID_TXF_9_TIERSTEMPS.getText())){
            addMessage(1,"PROBLEME \n-> Fichier des TIERS TEMPS non précisé -> 9. Parametres\n");
            bToutBon=false;
        }
        else {
            addMessage(1,"OK\n");
        }
        
        List<String> listHeures = new ArrayList<>(Arrays.asList("8h30","9h00","9h30","10h00","10h30","11h00","13h30","14h00","14h30","15h00","15h30","16h00","16h30"));
        List<String> listDuree  = new ArrayList<>(Arrays.asList("1h00","1h30","2h00","2h30","3h00"));
        List<String> listDeltaHeures = new ArrayList<>(Arrays.asList("-1h00","-0h40","-0h30","-0h20","0h00","0h20","0h30","0h40","1h00"));
        FXID_CCB_1_DEBUT_DS1.setItems(FXCollections.observableArrayList( listHeures));
        FXID_CCB_1_DEBUT_DS1.getSelectionModel().select(0);
        FXID_CCB_1_DUREE_DS1.setItems(FXCollections.observableArrayList( listDuree));
        FXID_CCB_1_DUREE_DS1.getSelectionModel().select(1);
        FXID_CCB_1_DEBUT_DS2.setItems(FXCollections.observableArrayList( listHeures));
        FXID_CCB_1_DEBUT_DS2.getSelectionModel().select(3);
        FXID_CCB_1_DUREE_DS2.setItems(FXCollections.observableArrayList( listDuree));
        FXID_CCB_1_DUREE_DS2.getSelectionModel().select(1);
        FXID_CCB_1_DELTATT_DS1.setItems(FXCollections.observableArrayList( listDeltaHeures));
        FXID_CCB_1_DELTATT_DS1.getSelectionModel().select(2);
        FXID_CCB_1_DELTATT_DS2.setItems(FXCollections.observableArrayList( listDeltaHeures));
        FXID_CCB_1_DELTATT_DS2.getSelectionModel().select(6);
    
        if (bToutBon) {
            addMessage(1,"----- INIT OK ---------\n");            
        }
        else {
            addMessage(1,"----- INIT PROBLEME ---\n");
        }
    }

    @FXML
    private void OnAction_TXF_1_RESULTAT(ActionEvent event) {
    }

    @FXML
    private void OnAction_BTN_1_RESULTAT(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            FXID_TXF_1_RESULTAT.setText(selectedFile.getPath());
        }
    }

    @FXML
    private void OnAction_BTN_1_GO(ActionEvent event) throws ClassNotFoundException, SQLException, IOException, InvalidFormatException, sExcelFileManagerException {
        addMessage(1,"----- GO --------\n");
        
        /***************************** 
         Lecture des valeurs choisies
        ******************************/ 
        
    // Date ------------------------
        LocalDate localDate = FXID_DPR_1_DATE.getValue();
        addMessage(1,"Date : "+localDate+"\n");
        m_DateSeed = localDate.getDayOfYear();
        
    // Fichier resultat ------------
        addMessage(1,"Fichier résultat : "+FXID_TXF_1_RESULTAT.getText()+"\n");
        
    // Liste des Salles sans doublons avec nombre d'etudiants par salle
        List<String> listTemp = new ArrayList<>();
        for (ChoiceBox<String> cb : m_ChoiceBoxSalles) {
            String s = cb.getSelectionModel().getSelectedItem() ;
            if (s.equalsIgnoreCase("nom")) {              
            }
            else {
                listTemp.add( cb.getSelectionModel().getSelectedItem() );
            }
        } 
        // Suppression des doublons
        Set<String> set = new HashSet<>();
        set.addAll(listTemp);
        List<String> listSalles = new ArrayList<>(set);
        String SalleAmenagee = FXID_CCB_1_SALLE00.getSelectionModel().getSelectedItem();
        Collections.sort(listSalles);
        addMessage(1,"Liste des salles :\n");
        for(String s: listSalles) {
            addMessage(1,s + "\n");
        }
        addMessage(1,"avec en plus "+ SalleAmenagee + " pour les Tiers Temps");
        
    // Nom des DS --------------
        m_DS1=FXID_CCB_1_DS1.getSelectionModel().getSelectedItem();
        if (m_DS1.contains("3")) m_Annee_DS1="3";
        if (m_DS1.contains("4")) m_Annee_DS1="4";
        if (m_DS1.contains("5")) m_Annee_DS1="5";
        m_DS2=FXID_CCB_1_DS2.getSelectionModel().getSelectedItem();
        if (m_DS2.contains("3")) m_Annee_DS2="3";
        if (m_DS2.contains("4")) m_Annee_DS2="4";
        if (m_DS2.contains("5")) m_Annee_DS2="5";
        addMessage(1,"DS1 : "+m_DS1+" pour les "+m_Annee_DS1+"IF\n");
        addMessage(1,"DS2 : "+m_DS2+" pour les "+m_Annee_DS2+"IF\n");
       
        
        
    /********************** 
              TRAITEMENTS
    ***********************/ 
        System.out.println("------- Début du traitement -------------");

        System.out.print("Connection ... ");
        String driver = "org.sqlite.JDBC";
        Class.forName(driver);
        String dbName = "sqlite_temp.db"; 
        String dbUrl = "jdbc:sqlite:" + dbName;
        Connection conn = DriverManager.getConnection(dbUrl);
        System.out.println("OK");

    //create table 
        System.out.print("CREATE TABLE Etudiants ... ");
        Statement stQuery = conn.createStatement();
        Statement stUpdate = conn.createStatement();
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS Etudiants");
        stUpdate.executeUpdate(   "CREATE TABLE Etudiants ("+ 
                            "Id             INTEGER PRIMARY KEY AUTOINCREMENT,"+ 
                            "Numero         CHAR(255)   NOT NULL,"+ 
                            "Nom            CHAR(255)   NOT NULL,"+ 
                            "Prenom         CHAR(64)    NOT NULL,"+
                            "Mec_code       CHAR(64)    NOT NULL,"+ 
                            "Annee          CHAR(64)    NOT NULL"+ 
                            ")");
        System.out.println("OK");

        System.out.print("CREATE TABLE EtudiantsTiersTempsTemp ... ");
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsTiersTempsTemp");
        stUpdate.executeUpdate(   "CREATE TABLE EtudiantsTiersTempsTemp ("+ 
                            "Id             INTEGER PRIMARY KEY AUTOINCREMENT,"+ 
                            "Numero         CHAR(255)   NOT NULL,"+ 
                            "Nom            CHAR(255)   NOT NULL,"+ 
                            "Prenom         CHAR(64)    NOT NULL"+
                            ")");
        System.out.println("OK");

        System.out.print("CREATE TABLE Salles ... ");
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS Salles");
        stUpdate.executeUpdate(   "CREATE TABLE Salles ("+ 
                            "Id             INTEGER PRIMARY KEY AUTOINCREMENT,"+ 
                            "Nom            CHAR(255)   NOT NULL,"+ 
                            "Capacite_DS    CHAR(8)     NOT NULL"+
                            ")");
        System.out.println("OK");

        
    //insert rows Etudiants
        String sql;
        PreparedStatement pstmt;
        System.out.print("INSERT INTO Etudiants ... ");
        sExcelFileManager sEFM_Etudiants = new sExcelFileManager(FXID_TXF_9_ETUDIANTS.getText());
        sEFM_Etudiants.setCurrentSheetFromSheetNumber(0);
        sql = "INSERT INTO Etudiants (Numero, Nom, Prenom, Mec_code, Annee) VALUES(?,?,?,?,?)";
        pstmt = conn.prepareStatement(sql);
        for (int ligne=1; ligne<sEFM_Etudiants.getCurrentSheetSize(); ligne++) {
            pstmt.setString(1, sEFM_Etudiants.getDataFromCurrentSheet(ligne, "Numero").toLowerCase());
            pstmt.setString(2, sEFM_Etudiants.getDataFromCurrentSheet(ligne, "Nom").toUpperCase());
            pstmt.setString(3, sEFM_Etudiants.getDataFromCurrentSheet(ligne, "Prenom").toLowerCase());
            pstmt.setString(4, sEFM_Etudiants.getDataFromCurrentSheet(ligne, "Mec_code").toUpperCase());
            pstmt.setString(5, sEFM_Etudiants.getDataFromCurrentSheet(ligne, "Annee").toLowerCase());
            pstmt.executeUpdate();            
        }
        System.out.println("OK");

    //insert rows Etudiants Echanges
        System.out.print("INSERT INTO Etudiants Echanges ... ");
        sExcelFileManager sEFM_Etudiants_Echanges = new sExcelFileManager(FXID_TXF_9_ETUDIANTS_ECHANGES.getText());
        sEFM_Etudiants_Echanges.setCurrentSheetFromSheetNumber(0);
        sql = "INSERT INTO Etudiants (Numero, Nom, Prenom, Mec_code, Annee) VALUES(?,?,?,?,?)";
        pstmt = conn.prepareStatement(sql);
        for (int ligne=1; ligne<sEFM_Etudiants_Echanges.getCurrentSheetSize(); ligne++) {
            pstmt.setString(1, sEFM_Etudiants_Echanges.getDataFromCurrentSheet(ligne, "Numero").toLowerCase());
            pstmt.setString(2, sEFM_Etudiants_Echanges.getDataFromCurrentSheet(ligne, "Nom").toUpperCase());
            pstmt.setString(3, sEFM_Etudiants_Echanges.getDataFromCurrentSheet(ligne, "Prenom").toLowerCase());
            pstmt.setString(4, sEFM_Etudiants_Echanges.getDataFromCurrentSheet(ligne, "Mec_code").toUpperCase());
            pstmt.setString(5, sEFM_Etudiants_Echanges.getDataFromCurrentSheet(ligne, "Annee").toLowerCase());
            pstmt.executeUpdate();            
        }
        System.out.println("OK");
        
    //insert rows TiersTemps
        System.out.print("INSERT INTO EtudiantsTiersTempsTemp ... ");
        sExcelFileManager sEFM_TiersTemps = new sExcelFileManager(FXID_TXF_9_TIERSTEMPS.getText());
        sEFM_TiersTemps.setCurrentSheetFromSheetNumber(0);
        sql = "INSERT INTO EtudiantsTiersTempsTemp (Numero, Nom, Prenom) VALUES(?,?,?)";
        pstmt = conn.prepareStatement(sql);
        for (int ligne=1; ligne<sEFM_TiersTemps.getCurrentSheetSize(); ligne++) {
            pstmt.setString(1, sEFM_TiersTemps.getDataFromCurrentSheet(ligne, "Numero").toLowerCase());
            pstmt.setString(2, sEFM_TiersTemps.getDataFromCurrentSheet(ligne, "Nom").toUpperCase());
            pstmt.setString(3, sEFM_TiersTemps.getDataFromCurrentSheet(ligne, "Prenom").toLowerCase());
            pstmt.executeUpdate();            
        }
        System.out.println("OK");

            
    // ajout du (T.M.) au nom des étudiants tiers temps
        ResultSet rs;
        String query;
        //query= "UPDATE Etudiants SET Nom = Nom || \" (T.M.)\" WHERE Numero IN (SELECT Numero FROM EtudiantsTiersTempsTemp";
        //stUpdate.executeUpdate(query);  
        query = "SELECT Numero FROM EtudiantsTiersTempsTemp";
        rs = stQuery.executeQuery(query);        
        try {
            while(rs.next()) {
                String numero = rs.getString(1);
                query = "UPDATE Etudiants SET Nom = Nom || \" (T.M.)\" WHERE Numero = '"+numero+"'";
                stUpdate.executeUpdate(query); 
            }
        }
        finally {
            rs.close();
        }
    
    

        
        
        
        
        //insert rows Salles
        System.out.print("INSERT INTO Salles ... ");
        sExcelFileManager sEFM_Salles = new sExcelFileManager(FXID_TXF_9_SALLES.getText());
        sEFM_TiersTemps.setCurrentSheetFromSheetNumber(0);
        sql = "INSERT INTO Salles (Nom, Capacite_DS) VALUES(?,?)";
        pstmt = conn.prepareStatement(sql);
        for (int ligne=1; ligne<sEFM_Salles.getCurrentSheetSize(); ligne++) {
            pstmt.setString(1, sEFM_Salles.getDataFromCurrentSheet(ligne, "Nom").toUpperCase());
            pstmt.setString(2, sEFM_Salles.getDataFromCurrentSheet(ligne, "Capacite_DS").toUpperCase());
            pstmt.executeUpdate();            
        }
        System.out.println("OK");
     
        
        
        
        // SELECT ...
        // Lister ceux qui : 
        //      - font le DS1 = EtudiantsDS1
        //      - ne font pas le DS1 = EtudiantsSansDS1
        //      - font le DS2 = EtudiantsDS2
        //      - ne font pas le DS2 = EtudiantsSansDS2
        //
        // Mélanger ceux qui font les deux DS et ne sont pas tiers temps = EtudiantsDS1DS2SansTiersTemps
        // Mettre en salle aménagés ceux qui : = EtudiantsAmenages
        //      - sont tiers temps = EtudiantsTiersTemps
        //      - font le DS1 et pas le DS2 = EtudiantsDS1SansDS2
        //      - font le DS2 et pas le DS1 = EtudiantsDS2SansDS1
        //
        
    //select EtudiantsTiersTemps 
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsTiersTemps");
        query = "CREATE TABLE EtudiantsTiersTemps AS "+
                "SELECT Numero, Nom, Prenom, Annee, Mec_code FROM Etudiants "+
                "WHERE Numero IN (SELECT Numero FROM EtudiantsTiersTempsTemp)";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);  
        
    //select Etudiants qui font le DS1 
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsDS1");
        query = "CREATE TABLE EtudiantsDS1 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee, Mec_code FROM Etudiants WHERE Mec_code = '"+m_DS1+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);  
        System.out.println("- query OK -");

        
        
    //select Etudiants qui font le DS2 
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsDS2");
        query = "CREATE TABLE EtudiantsDS2 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee, Mec_code FROM Etudiants WHERE Mec_code = '"+m_DS2+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

//        try {
//           query ="SELECT Numero, Nom, Prenom, Annee FROM EtudiantsDS2";
//           rs = st.executeQuery(query);
//           int index=1;
//           while(rs.next()) {
//                  String Numero = rs.getString(1);
//                  String Nom = rs.getString(2).toUpperCase();
//                  String Prenom = rs.getString(3);
//                  String Annee = rs.getString(4);
//                  System.out.println(index +"\t"+Numero+" "+Annee+" : "+Nom +" "+Prenom);
//                  index++;
//            }
//        } 
//        finally {
//            rs.close();
//        }
        System.out.println("- query OK -");

        
        
    //select Etudiants de nIF qui ne font pas le DS1 
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsSansDS1");
        query = "CREATE TABLE EtudiantsSansDS1 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Annee = '"+m_Annee_DS1+"'" +
                " EXCEPT "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Mec_code = '"+m_DS1+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        System.out.println("- query OK -");

        
    //select Etudiants de nIF qui ne font pas le DS2 
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsSansDS2");
        query = "CREATE TABLE EtudiantsSansDS2 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Annee = '"+m_Annee_DS2+"'" +
                " EXCEPT "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Mec_code = '"+m_DS2+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        System.out.println("- query OK -");
        
    


        
    //select Etudiants qui font le DS1 et pas le DS2 = EtudiantsDS1SansDS2
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsDS1SansDS2");
        query = "CREATE TABLE EtudiantsDS1SansDS2 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Mec_code = '"+m_DS1+"'"+
                " EXCEPT "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Mec_code = '"+m_DS2+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        System.out.println("- query OK -");
        
                
    //select Etudiants qui font le DS2 et pas le DS1 = EtudiantsDS2SansDS1
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsDS2SansDS1");
        query = "CREATE TABLE EtudiantsDS2SansDS1 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Mec_code = '"+m_DS2+"'"+
                " EXCEPT "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Etudiants WHERE Mec_code = '"+m_DS1+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        System.out.println("- query OK -");
        
        
        
        
        
        
        
    //select Etudiants qui font le DS1 et le DS2 sans tiers temps = EtudiantsDS1DS2SansTiersTemps
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsDS1DS2SansTiersTemps");
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS Temp");
        query = "CREATE TABLE Temp AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsDS1"+
                " INTERSECT "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsDS2"+
                " EXCEPT "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsTiersTemps";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        query = "CREATE TABLE EtudiantsDS1DS2SansTiersTemps AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Temp"+
                " ORDER BY Nom, Prenom, Numero";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        
        query = "ALTER TABLE EtudiantsDS1DS2SansTiersTemps ADD COLUMN Salle CHAR(32)";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        
        query = "SELECT COUNT(*) FROM EtudiantsDS1DS2SansTiersTemps";
        rs = stQuery.executeQuery(query);
        int NbDS1DS2=Integer.parseInt(rs.getString(1));
        
        query = "SELECT Numero, Annee, Nom, Prenom FROM EtudiantsDS1DS2SansTiersTemps";
        rs = stQuery.executeQuery(query);
        try {
           PrintRS(rs, NbDS1DS2, 4);
        } 
        finally {
            rs.close();
        }
        System.out.println("- query OK -");
        
      
        
        
        
        
        
    //select Etudiants Aménagés DS1 = TiersTemps DS1 + DS1sansDS2 = EtudiantsAmenagesDS1
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsAmenagesDS1");
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS Temp");
        query = "CREATE TABLE Temp AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsDS1SansDS2"+
                " UNION "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsTiersTemps WHERE Mec_code='"+m_DS1+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        query = "CREATE TABLE EtudiantsAmenagesDS1 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Temp"+
                " ORDER BY Nom, Prenom, Numero";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        query = "ALTER TABLE EtudiantsAmenagesDS1 ADD COLUMN Salle CHAR(32)";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);
        
        query = "SELECT COUNT(*) FROM EtudiantsAmenagesDS1";
        rs = stQuery.executeQuery(query);
        int NbAmenDS1=Integer.parseInt(rs.getString(1));
        
        query = "SELECT Numero, Annee, Nom, Prenom FROM EtudiantsAmenagesDS1";
        rs = stQuery.executeQuery(query);
        try {
           PrintRS(rs, NbAmenDS1, 4);
        } 
        finally {
            rs.close();
        }

        
        
        
    //select Etudiants Aménagés DS2 = TiersTemps DS2 + DS2sansDS1 = EtudiantsAmenagesDS2
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS EtudiantsAmenagesDS2");
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS Temp");
        query = "CREATE TABLE Temp AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsDS2SansDS1"+
                " UNION "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM EtudiantsTiersTemps WHERE Mec_code='"+m_DS2+"'";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        query = "CREATE TABLE EtudiantsAmenagesDS2 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee FROM Temp"+
                " ORDER BY Nom, Prenom, Numero";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        query = "ALTER TABLE EtudiantsAmenagesDS2 ADD COLUMN Salle CHAR(32)";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        query = "SELECT COUNT(*) FROM EtudiantsAmenagesDS2";
        rs = stQuery.executeQuery(query);
        int NbAmenDS2=Integer.parseInt(rs.getString(1));

        query = "SELECT Numero, Nom, Prenom, Annee FROM EtudiantsAmenagesDS2";
        rs = stQuery.executeQuery(query);
        try {
           PrintRS(rs, NbAmenDS2, 4);
        } 
        finally {
            rs.close();
        }
            
        
    // Affectation des salles par permutation
    
        int NbEtudiantsMax = Math.max(NbDS1DS2+NbAmenDS1, NbDS1DS2+NbAmenDS2);
        int NbSalles=listSalles.size()+1; // +1 pour la salle des aménagés
        int NbPlacesAmenagees=Integer.parseInt(SalleAmenagee.substring(SalleAmenagee.length()-3, SalleAmenagee.length()).trim());
        int NbPlacesTotal=NbPlacesAmenagees;
        int n;
        Integer[] tabCap = new Integer[listSalles.size()]; // Tableau de capacité des salles
        String[] tabSalle= new String[listSalles.size()];  // Tableau des noms des salles
        for (int i=0; i<listSalles.size(); i++) {
            String s=listSalles.get(i);
            tabSalle[i]=s.substring(0, s.length()-5).trim(); //récupération des noms
            n=Integer.parseInt(s.substring(s.length()-3, s.length()).trim()); // récup. des capacités sur les 3 derniers caractères
            tabCap[i]=n;
            NbPlacesTotal+=n;
        }
        int NbPlacesLibres = NbPlacesTotal - NbEtudiantsMax;
        int NbPlacesLibresParSalles = NbPlacesLibres / NbSalles; // Arrondi ? Calcul entier
        String[] tabPlacesAffectees = new String[NbDS1DS2];
        int k=0;
        for (int i=0; i<listSalles.size(); i++) {
            for(int j=0;j<tabCap[i]-NbPlacesLibresParSalles; j++) {
                tabPlacesAffectees[k]= tabSalle[i];
                k++;
            }
        }
        // On complete les places à permuter par quelques places dans la salle des aménagés
        System.out.println("Nombre de places a permuter : "+ NbDS1DS2);
        System.out.println("Nombre de places affectées  : "+ k);
        int reste = NbDS1DS2-k;
        for (int j=0; j<reste;j++) {
            String s = SalleAmenagee;
            tabPlacesAffectees[k]=s.substring(0, s.length()-5).trim(); // récupération du nom
            k++;
        }
        
        System.out.println("------ Affectation des salles -----------");
        System.out.println("Nombre de places a permuter : "+ NbDS1DS2);
        System.out.println("Nombre de places affectées  : "+ k);
        for (int i=0; i<NbDS1DS2; i++) {
            System.out.println(i+" : "+tabPlacesAffectees[i]);       
        }
        
        System.out.println("----- Permutations ------");       
        
        // Permutations
        Random rand = new Random(m_DateSeed); 
        for (int i=0;i<NbDS1DS2-1; i++) {
            int nombreAleatoire = rand.nextInt(NbDS1DS2 - i-1) + i+1;
            String temp = tabPlacesAffectees[i];
            tabPlacesAffectees[i]=tabPlacesAffectees[nombreAleatoire];
            tabPlacesAffectees[nombreAleatoire]=temp;
        }
        for (int i=0; i<NbDS1DS2; i++) {
            System.out.println(i+" : "+tabPlacesAffectees[i]);  
        }
     
    // Affectation des salles dans les tables de la base
        
        query = "SELECT Numero FROM EtudiantsDS1DS2SansTiersTemps";
        System.out.println(query);
        rs = stQuery.executeQuery(query);        
        try {
            int i=0;
            while(rs.next()) {
                String numero = rs.getString(1);
                System.out.println("UPDATE : "+numero+"  "+tabPlacesAffectees[i]);
                query = "UPDATE EtudiantsDS1DS2SansTiersTemps SET Salle='"+tabPlacesAffectees[i]+"' WHERE Numero = '"+numero+"'";
                stUpdate.executeUpdate(query); 
                i++;
            }
        }
        finally {
            rs.close();
        }
        System.out.println("- query OK -");
  
        
        
        // Affectation de la salle aménagée
        String NomSalleAmenagee = SalleAmenagee.substring(0, SalleAmenagee.length()-5).trim(); // récupération du nom
        query = "SELECT Numero FROM EtudiantsAmenagesDS1";
        rs = stQuery.executeQuery(query);        
        try {
            int i=0;
            while(rs.next()) {
                query = "UPDATE EtudiantsAmenagesDS1 SET Salle='"+NomSalleAmenagee+"' WHERE Numero = '"+rs.getString(1)+"'";
                stUpdate.executeUpdate(query);            
            }
        }
        finally {
            rs.close();
        }
        System.out.println("- query OK -");

        query = "SELECT Numero FROM EtudiantsAmenagesDS2";
        System.out.println(query);
        rs = stQuery.executeQuery(query);        
        try {
            int i=0;
            while(rs.next()) {
                String numero = rs.getString(1);
                System.out.println("UPDATE : "+numero);
                query = "UPDATE EtudiantsAmenagesDS2 SET Salle='"+NomSalleAmenagee+"' WHERE Numero = '"+numero+"'";
                stUpdate.executeUpdate(query);            
            }
        }
        finally {
            rs.close();
        }
        System.out.println("- query OK -");
      
        
        
    // LISTES
        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS ListeDS1");
        query = "CREATE TABLE ListeDS1 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee, Salle FROM EtudiantsDS1DS2SansTiersTemps"+
                " UNION "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee, Salle FROM EtudiantsAmenagesDS1";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);

        stUpdate.executeUpdate(   "DROP TABLE IF EXISTS ListeDS2");
        query = "CREATE TABLE ListeDS2 AS "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee, Salle FROM EtudiantsDS1DS2SansTiersTemps"+
                " UNION "+
                "SELECT DISTINCT Numero, Nom, Prenom, Annee, Salle FROM EtudiantsAmenagesDS2";
        System.out.println("- query -----------------------");
        System.out.println(query);
        stUpdate.executeUpdate(query);


 // Extraction des résultats et enregistrements dans le fichier EXCEL   
        sExcelFileManager sEFMout = new sExcelFileManager();
        sEFMout.setSaveFileName("./test.xlsx");
        
    // Liste alphabetique des Etudiants du DS1
        query = "SELECT Numero, Nom, Prenom, Annee, Salle FROM ListeDS1 ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 5);
           sEFMout.newSheetFromResultSet(m_DS1, rs);
        } 
        finally {
            rs.close();
        }
        
    // Liste alpha des Etudiants sans DS1 : EtudiantsSansDS1    
        query = "SELECT Numero, Nom, Prenom, Annee FROM EtudiantsSansDS1 ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 5);
           sEFMout.newSheetFromResultSet("Sans "+m_DS1, rs);
        } 
        finally {
            rs.close();
        }
        
        
        
        
            
    // Liste alphabetique des Etudiants du DS2
        query = "SELECT Numero, Nom, Prenom, Annee, Salle FROM ListeDS2 ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 5);
           sEFMout.newSheetFromResultSet(m_DS2, rs);
        } 
        finally {
            rs.close();
        }
        
    // Liste alpha des Etudiants sans DS2 : EtudiantsSansDS2    
        query = "SELECT Numero, Nom, Prenom, Annee FROM EtudiantsSansDS2 ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 5);
           sEFMout.newSheetFromResultSet("Sans "+m_DS2, rs);
        } 
        finally {
            rs.close();
        }
            
        
    // Liste des Etudiants du DS1 triée par salle
        query = "SELECT Numero, Nom, Prenom, Annee, Salle FROM ListeDS1 ORDER BY Salle, Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 5);
           sEFMout.newSheetFromResultSet(m_DS1+" tri par salle" , rs);
        } 
        finally {
            rs.close();
        }
        
    // Liste des Etudiants du DS2 triée par salle
        query = "SELECT Numero, Nom, Prenom, Annee, Salle FROM ListeDS2 ORDER BY Salle, Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 5);
           sEFMout.newSheetFromResultSet(m_DS2+" tri par salle" , rs);
        } 
        finally {
            rs.close();
        }
        
    // Liste alpha des Etudiants du DS1 sans DS2
        query = "SELECT Numero, Nom, Prenom, Annee FROM EtudiantsDS1SansDS2 ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 4);
           sEFMout.newSheetFromResultSet(m_DS1+" sans "+m_DS2 , rs);
        } 
        finally {
            rs.close();
        }
        
    // Liste alpha des Etudiants du DS2 sans DS1
        query = "SELECT Numero, Nom, Prenom, Annee FROM EtudiantsDS2SansDS1 ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           //PrintRS(rs, -1, 4);
           sEFMout.newSheetFromResultSet(m_DS2+" sans "+m_DS1 , rs);
        } 
        finally {
            rs.close();
        }


        
        // Liste des noms de salles : SalleAmenagee et les listSalles.size() salles  : listSalles.get(i)
        sDSData DSdata = new sDSData();
        
    // DS1
        DSdata.setNom(m_DS1);
        DSdata.setDate(localDate);
        DSdata.setDebut(FXID_CCB_1_DEBUT_DS1.getSelectionModel().getSelectedItem());
        DSdata.setDuree(FXID_CCB_1_DUREE_DS1.getSelectionModel().getSelectedItem());
        DSdata.setDeltaTT(FXID_CCB_1_DELTATT_DS1.getSelectionModel().getSelectedItem());
        
        // Requete pour savoir si il y a des tiers temps dans le DS1
        query = "SELECT COUNT(*) FROM (SELECT A.Numero FROM ListeDS1 A, EtudiantsTiersTemps B WHERE A.Numero = B.Numero)";
        rs = stQuery.executeQuery(query);
        if (Integer.parseInt(rs.getString(1))>0) {
            DSdata.setTiersTemps(true);
        }

        // SalleAmenagee
        String s=SalleAmenagee;
        String salle = s.substring(0, s.length()-5).trim();    
        DSdata.setSalle(salle);
        query = "SELECT Numero, Nom, Prenom FROM ListeDS1 WHERE Salle='"+salle+"' ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           sEFMout.newSheetFromResultSet_DS(m_DS1+" - "+salle , rs, DSdata);
        } 
        finally {
            rs.close();
        }
                
        
        for (int i=0; i<listSalles.size(); i++) {
            s=listSalles.get(i);
            salle = s.substring(0, s.length()-5).trim();    
            DSdata.setSalle(salle);
            query = "SELECT Numero, Nom, Prenom FROM ListeDS1 WHERE Salle='"+salle+"' ORDER BY Nom, Prenom, Numero ASC";
            rs = stQuery.executeQuery(query);
            try {
               //PrintRS(rs, -1, 4);
               sEFMout.newSheetFromResultSet_DS(m_DS1+" - "+salle , rs, DSdata);
            } 
            finally {
                rs.close();
            }
        }

        
        
    // DS2
        DSdata.setNom(m_DS2);
        DSdata.setDate(localDate);
        DSdata.setDebut(FXID_CCB_1_DEBUT_DS2.getSelectionModel().getSelectedItem());
        DSdata.setDuree(FXID_CCB_1_DUREE_DS2.getSelectionModel().getSelectedItem());
        DSdata.setDeltaTT(FXID_CCB_1_DELTATT_DS2.getSelectionModel().getSelectedItem());

        // Requete pour savoir si il y a des tiers temps dans le DS2
        query = "SELECT COUNT(*) FROM (SELECT A.Numero FROM ListeDS2 A, EtudiantsTiersTemps B WHERE A.Numero = B.Numero)";
        rs = stQuery.executeQuery(query);
        if (Integer.parseInt(rs.getString(1))>0) {
            DSdata.setTiersTemps(true);
        }
        
        // SalleAmenagee
        s=SalleAmenagee;
        salle = s.substring(0, s.length()-5).trim();    
        DSdata.setSalle(salle);
        query = "SELECT Numero, Nom, Prenom FROM ListeDS2 WHERE Salle='"+salle+"' ORDER BY Nom, Prenom, Numero ASC";
        rs = stQuery.executeQuery(query);
        try {
           sEFMout.newSheetFromResultSet_DS(m_DS2+" - "+salle , rs, DSdata);
        } 
        finally {
            rs.close();
        }
                
        
        for (int i=0; i<listSalles.size(); i++) {
            s=listSalles.get(i);
            salle = s.substring(0, s.length()-5).trim();    
            DSdata.setSalle(salle);
            query = "SELECT Numero, Nom, Prenom FROM ListeDS2 WHERE Salle='"+salle+"' ORDER BY Nom, Prenom, Numero ASC";
            rs = stQuery.executeQuery(query);
            try {
               //PrintRS(rs, -1, 4);
               sEFMout.newSheetFromResultSet_DS(m_DS2+" - "+salle , rs, DSdata);
            } 
            finally {
                rs.close();
            }
        }

        
        
        
    
        String path = "/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/";
        sEFMout.saveAs(path+localDate.toString()+"_"+m_DS1+"_"+m_DS2+".xlsx");



        
        System.out.println("Opened database successfully");    
    }
    
    private void PrintRS(ResultSet rs, int nbligne, int NbCol) throws SQLException {
        int index=0;
        
        String NbLigne= Integer.toString(nbligne);
        if (nbligne<0) {
            NbLigne=" ? ";
        }
        while(rs.next()) {
            index++;
            System.out.print(index+"/"+NbLigne+"\t: ");
            for (int c=1; c<=NbCol;c++){
                String S = rs.getString(c);
                System.out.print(S+" ");
            }
            System.out.println("");
        }
        System.out.println("- "+NbLigne+" réponses - ");
    }
        
        
    

    @FXML
    private void OnAction_BTN_QUIT(ActionEvent event) {
        // get a handle to the stage
        Stage stage = (Stage) FXID_BTN_QUIT.getScene().getWindow();
        // do what you have to do
        stage.close();
    }

    @FXML
    private void OnAction_BTN_9_ETUDIANTS_ECHANGES(ActionEvent event) {
    }

    @FXML
    private void OnAction_BTN_2_INIT(ActionEvent event) throws ClassNotFoundException, SQLException, IOException, InvalidFormatException, sExcelFileManagerException {
        
        /*************************************************
         * Initialisation des données pour le calcul des heures
         * Les fichiers sources des données sont pris aux adresses :
         * 
         * https://servif-cocktail.insa-lyon.fr/EdT/3IF.csv
         * https://servif-cocktail.insa-lyon.fr/EdT/4IF.csv
         * https://servif-cocktail.insa-lyon.fr/EdT/5IF.csv
         * 
         * ce sont des extractions directes des Emplois du temps SuperPlan
         * 
         */
        
        addMessage(2,"INITIALISATION ... ");
        
        
        /***************************** 
         Lecture des valeurs choisies
        ******************************/ 
        
        
    // Fichier resultat ------------
    //    addMessage(2,"Fichier résultat : "+FXID_TXF_1_RESULTAT.getText()+"\n");
        
       
        
        
    /********************** 
              TRAITEMENTS
    ***********************/ 
        System.out.println("------- Début du traitement -------------");
        
        String path = "/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/";

    // Création de la base
        sSQLiteManager dbDeclaration = new sSQLiteManager("sqlite_Declaration.db");
        dbDeclaration.createTableYannDropIfExists("tableCreneaux");
//        dbDeclaration.createTableMaquetteDropIfExists("tableMaquette");
//        System.out.println("Creation tables OK");
    //insert rows tableCreneaux
        dbDeclaration.addTableYann("tableCreneaux", FXID_TXF_9_CRENEAUX_3IF.getText(), "3IF");
        System.out.println("3IF ajouté.");
        dbDeclaration.addTableYann("tableCreneaux", FXID_TXF_9_CRENEAUX_4IF.getText(), "4IF");
        System.out.println("4IF ajouté.");
        dbDeclaration.addTableYann("tableCreneaux", FXID_TXF_9_CRENEAUX_5IF.getText(), "5IF");
        System.out.println("5IF ajouté.");
        dbDeclaration.convertExcelFileIntoTable(path + "config-Maquette-2017-2018.xlsx", "tableMaquette");
        dbDeclaration.convertExcelFileIntoTable(path + "config-Enseignant-2017-2018.xlsx", "tableEnseignant");
        dbDeclaration.convertExcelFileIntoTable(path + "config-CoefTP-2017-2018.xlsx", "tableCoefTP");
//        dbDeclaration.addTableMaquette("tableMaquette", m_Fichier_Maquette); //FXID_TXF_9_CRENEAUX_5IF.getText());
        System.out.println("tableMaquette et tableEnseignant et tableCoefTP ajoutées.");

        //dbDeclaration.convertExcelFileIntoTable(path+"OR_EDT-2018.01.14.xlsx", "tableOR_EDT");

    
    //  
        dbDeclaration.complementTableCreneauxYann("tableCreneaux", "tableMaquette");
        addMessage(2," Complement Table OK.");
    
        String sql = "SELECT DISTINCT Numero, Enseignant, Nom, Prenom, EC_Code, Module, Type, DureeMinutes, Jour, Heure_Debut, Heure_Fin, Annee "+
                        "FROM tableCreneaux ORDER BY Nom, Prenom, EC_Code, Type";
        ResultSet res =  dbDeclaration.querySQL(sql);
        sExcelFileManager sEFM = new sExcelFileManager();
        sEFM.newSheetFromResultSet("Declaration Brute", res);
        sEFM.saveAs(path+"testDeclarationBrute.xlsx");
    
        sql = "DROP TABLE IF EXISTS tableDeclarationTemp";
        dbDeclaration.updateSQL(sql);
        sql = "CREATE TABLE tableDeclarationTemp AS " +
              "SELECT Numero, EC_Code, Nom, Prenom, Type, SUM(DureeMinutes) AS sumDureeMinutes, SUM(DureeMinutes)/60.0 AS sumDureeHeure "
                + " FROM tableCreneaux "
                + " WHERE EC_Code IS NOT NULL "    
                + " GROUP BY Numero, EC_Code, Type "
                + " ORDER BY Nom, Prenom, EC_Code, Type; ";
        dbDeclaration.updateSQL(sql);


        /******************
        * Création de la table tableDeclarationEffectif qui contient les déclarations 
        * des heures reélles EFFECTIVES en Heures, AVANT application des coéfficients ...
        ********************/
        sql = "DROP TABLE IF EXISTS tableDeclarationEffectif";
        dbDeclaration.updateSQL(sql);
        sql = "CREATE TABLE tableDeclarationEffectif AS "+
              "SELECT DISTINCT Numero, Nom, Prenom, EC_Code FROM tableDeclarationTemp "+
              "WHERE (Nom IS NOT NULL) AND (EC_Code IS NOT NULL) "+
              "ORDER BY Nom, Prenom, Numero, EC_Code";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN CM CHAR(8)";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN TD CHAR(8)";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN TP CHAR(8)";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN Projet CHAR(8)";
        dbDeclaration.updateSQL(sql);
        
        sql="UPDATE tableDeclarationEffectif " +
            "SET CM = (SELECT D.sumDureeHeure "+
            "           FROM tableDeclarationTemp D " +
            "           WHERE D.Type = 'CM' AND tableDeclarationEffectif.Numero = D.Numero AND tableDeclarationEffectif.EC_Code = D.EC_Code), "+
            "    TD = (SELECT D.sumDureeHeure "+
            "           FROM tableDeclarationTemp D " +
            "           WHERE D.Type = 'TD' AND tableDeclarationEffectif.Numero = D.Numero AND tableDeclarationEffectif.EC_Code = D.EC_Code), "+
            "    TP = (SELECT D.sumDureeHeure "+
            "           FROM tableDeclarationTemp D " +
            "           WHERE D.Type = 'TP' AND tableDeclarationEffectif.Numero = D.Numero AND tableDeclarationEffectif.EC_Code = D.EC_Code) ";
        dbDeclaration.updateSQL(sql);
        
        // Sauvegarde de la table dans un fichier EXCEL
        sql = "SELECT Numero AS NumIndividu, EC_Code AS Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationEffectif ORDER BY Nom, Prenom, Ects";
        res =  dbDeclaration.querySQL(sql);
        sEFM = new sExcelFileManager();
        sEFM.newSheetFromResultSet("DeclarationEffectif", res);
        sEFM.saveAs(path+"tableDeclarationEffectif.xlsx");  
     
        /******************
        * Création de la table tableDeclarationPonderee qui contient les déclarations 
        * des heures reelles EFFECTIVES en Heures PONDEREES par les coéfficients ADHOC...
        ********************/
        sql = "DROP TABLE IF EXISTS tableDeclarationPonderee";
        dbDeclaration.updateSQL(sql);
        sql = "CREATE TABLE tableDeclarationPonderee AS "+
              "SELECT Numero AS NumIndividu, EC_Code AS Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationEffectif ORDER BY Nom, Prenom, Ects";
        dbDeclaration.updateSQL(sql);
        // Pondération des TD
        sql="UPDATE tableDeclarationPonderee " +
            "SET TD = TD * 10.0/8.0 "+
            "WHERE TD IS NOT NULL;";
        dbDeclaration.updateSQL(sql);
        // Pondération des TP
        sql="UPDATE tableDeclarationPonderee " +
            "SET TP = TP * 24.0/16.0 "+
            "WHERE (TP IS NOT NULL) "
            + "AND ((SELECT CoefTP FROM tableCoefTP WHERE tableCoefTP.EC_Code = tableDeclarationPonderee.Ects)='24.0');";
        dbDeclaration.updateSQL(sql);
        sql="UPDATE tableDeclarationPonderee " +
            "SET TP = TP * 28.5/16.0 "+
            "WHERE (TP IS NOT NULL) "
            + "AND ((SELECT CoefTP FROM tableCoefTP WHERE tableCoefTP.EC_Code = tableDeclarationPonderee.Ects)='28.5');";
        dbDeclaration.updateSQL(sql);

        // Sauvegarde de la table dans un fichier EXCEL
        sql = "SELECT NumIndividu, Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationPonderee " +
              "WHERE ((SELECT Statut FROM tableEnseignant WHERE tableEnseignant.Numero = tableDeclarationPonderee.NumIndividu)='PERM') "+  
              "ORDER BY Nom, Prenom, Ects;";
        res =  dbDeclaration.querySQL(sql);
        sEFM = new sExcelFileManager();
        sEFM.newSheetFromResultSet("Déclaration Permanents", res);
        sEFM.saveAs(path+"tableDeclarationPondereePerm.xlsx");  
        
        // Sauvegarde de la table dans un fichier EXCEL
        sql = "SELECT NumIndividu, Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationPonderee " +
              "WHERE ((SELECT Statut FROM tableEnseignant WHERE tableEnseignant.Numero = tableDeclarationPonderee.NumIndividu)='VACA') "+  
              "ORDER BY Nom, Prenom, Ects;";
        res =  dbDeclaration.querySQL(sql);
        sEFM = new sExcelFileManager();
        sEFM.newSheetFromResultSet("Déclaration Vacataires", res);
        sEFM.saveAs(path+"tableDeclarationPondereeVaca.xlsx");  
        
        
        addMessage(2,"FIN OK.");

    }

    @FXML
    private void OnAction_BTN_9_CRENEAUX_3IF(ActionEvent event) {
    }

    @FXML
    private void OnAction_BTN_9_CRENEAUX_4IF(ActionEvent event) {
    }

    @FXML
    private void OnAction_BTN_9_CRENEAUX_5IF(ActionEvent event) {
    }

    @FXML
    private void OnAction_BTN_2_CALCUL_HEURES(ActionEvent event) throws ClassNotFoundException, SQLException, IOException, InvalidFormatException, sExcelFileManagerException {
        /*************************************************
         * Initialisation des données pour le calcul des heures
         * Les fichiers sources des données sont pris aux adresses :
         * 
         * https://servif-cocktail.insa-lyon.fr/EdT/3IF.csv
         * https://servif-cocktail.insa-lyon.fr/EdT/4IF.csv
         * https://servif-cocktail.insa-lyon.fr/EdT/5IF.csv
         * 
         * ce sont des extractions directes des Emplois du temps SuperPlan
         * 
         */
        
        addMessage(2,"CALCUL DES HEURES ... ");
        
        
    /********************** 
              TRAITEMENTS
    ***********************/ 
        System.out.println("------- Début du traitement -------------");
        
        String path = "/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/";

    // Création de la base
        sSQLiteManager dbDeclaration = new sSQLiteManager("sqlite_Declaration.db");
        dbDeclaration.createTableYannDropIfExists("tableCreneaux");
//        dbDeclaration.createTableMaquetteDropIfExists("tableMaquette");
//        System.out.println("Creation tables OK");
    //insert rows tableCreneaux

        dbDeclaration.addTableYann("tableCreneaux", FXID_TXF_9_CRENEAUX_3IF.getText(), "3IF");
        System.out.println("3IF ajouté.");
        dbDeclaration.addTableYann("tableCreneaux", FXID_TXF_9_CRENEAUX_4IF.getText(), "4IF");
        System.out.println("4IF ajouté.");
        dbDeclaration.addTableYann("tableCreneaux", FXID_TXF_9_CRENEAUX_5IF.getText(), "5IF");
        System.out.println("5IF ajouté.");
        dbDeclaration.createAndFillTableFromExcelSheet("tableMaquette", path + "config-EffectIF-2017-2018.xlsx", "Maquette");
        dbDeclaration.createAndFillTableFromExcelSheet("tableEnseignant", path + "config-EffectIF-2017-2018.xlsx", "Enseignant");
        dbDeclaration.createAndFillTableFromExcelSheet("tableCoefTP", path + "config-EffectIF-2017-2018.xlsx", "CoefTP");
//        dbDeclaration.addTableMaquette("tableMaquette", m_Fichier_Maquette); //FXID_TXF_9_CRENEAUX_5IF.getText());
        System.out.println("tableMaquette et tableEnseignant et tableCoefTP ajoutées.");
    
    //  
        dbDeclaration.complementTableCreneauxYann("tableCreneaux", "tableMaquette");
        addMessage(2," Complement Table OK.");
    
        String sql = "SELECT DISTINCT Numero, Enseignant, Nom, Prenom, EC_Code, Module, Type, DureeMinutes, Jour, Heure_Debut, Heure_Fin, Annee "+
                        "FROM tableCreneaux ORDER BY Nom, Prenom, EC_Code, Type";
        ResultSet res =  dbDeclaration.querySQL(sql);
//        Workbook workbook = dbDeclaration.appendResultSetAsSheetToExcelFile(res, "Declaration Brute", path+"testDeclarationBrute.xlsx");
//        dbDeclaration.saveWorkbookToExcelFile(workbook, path+"testDeclarationBrute.xlsx");
        sExcelFileManager sEFM = new sExcelFileManager();
        sEFM.newSheetFromResultSet("Declaration Brute", res);
//        sEFM.setWorkbook(workbook);
        sEFM.saveAs(path+"testDeclarationBrute.xlsx");
        System.out.println("-----Déclaration Brute OK.");
        
        sql = "DROP TABLE IF EXISTS tableDeclarationTemp";
        dbDeclaration.updateSQL(sql);
        sql = "CREATE TABLE tableDeclarationTemp AS " +
              "SELECT Numero, EC_Code, Nom, Prenom, Type, SUM(DureeMinutes) AS sumDureeMinutes, SUM(DureeMinutes)/60.0 AS sumDureeHeure "
                + " FROM tableCreneaux "
                + " WHERE EC_Code IS NOT NULL "    
                + " GROUP BY Numero, EC_Code, Type "
                + " ORDER BY Nom, Prenom, EC_Code, Type; ";
        dbDeclaration.updateSQL(sql);


        /******************
        * Création de la table tableDeclarationEffectif qui contient les déclarations 
        * des heures reélles EFFECTIVES en Heures, AVANT application des coéfficients ...
        ********************/
        sql = "DROP TABLE IF EXISTS tableDeclarationEffectif";
        dbDeclaration.updateSQL(sql);
        sql = "CREATE TABLE tableDeclarationEffectif AS "+
              "SELECT DISTINCT Numero, Nom, Prenom, EC_Code FROM tableDeclarationTemp "+
              "WHERE (Nom IS NOT NULL) AND (EC_Code IS NOT NULL) "+
              "ORDER BY Nom, Prenom, Numero, EC_Code";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN CM CHAR(8)";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN TD CHAR(8)";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN TP CHAR(8)";
        dbDeclaration.updateSQL(sql);
        sql = "ALTER TABLE tableDeclarationEffectif ADD COLUMN Projet CHAR(8)";
        dbDeclaration.updateSQL(sql);
        
        sql="UPDATE tableDeclarationEffectif " +
            "SET CM = (SELECT D.sumDureeHeure "+
            "           FROM tableDeclarationTemp D " +
            "           WHERE D.Type = 'CM' AND tableDeclarationEffectif.Numero = D.Numero AND tableDeclarationEffectif.EC_Code = D.EC_Code), "+
            "    TD = (SELECT D.sumDureeHeure "+
            "           FROM tableDeclarationTemp D " +
            "           WHERE D.Type = 'TD' AND tableDeclarationEffectif.Numero = D.Numero AND tableDeclarationEffectif.EC_Code = D.EC_Code), "+
            "    TP = (SELECT D.sumDureeHeure "+
            "           FROM tableDeclarationTemp D " +
            "           WHERE D.Type = 'TP' AND tableDeclarationEffectif.Numero = D.Numero AND tableDeclarationEffectif.EC_Code = D.EC_Code) ";
        dbDeclaration.updateSQL(sql);
        
        // Sauvegarde de la table dans un fichier EXCEL
        sql = "SELECT Numero AS NumIndividu, EC_Code AS Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationEffectif ORDER BY Nom, Prenom, Ects";
        res =  dbDeclaration.querySQL(sql);
        dbDeclaration.appendResultSetAsSheetToExcelFile(res, "DeclarationEffectif", path+"tableDeclarationEffectif.xlsx");
        //sEFM = new sExcelFileManager();
        //sEFM.newSheetFromResultSet("DeclarationEffectif", res);
        //sEFM.saveAs(path+"tableDeclarationEffectif.xlsx");  
     
        /******************
        * Création de la table tableDeclarationPonderee qui contient les déclarations 
        * des heures reelles EFFECTIVES en Heures PONDEREES par les coéfficients ADHOC...
        ********************/
        sql = "DROP TABLE IF EXISTS tableDeclarationPonderee";
        dbDeclaration.updateSQL(sql);
        sql = "CREATE TABLE tableDeclarationPonderee AS "+
              "SELECT Numero AS NumIndividu, EC_Code AS Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationEffectif ORDER BY Nom, Prenom, Ects";
        dbDeclaration.updateSQL(sql);
        // Pondération des TD
        sql="UPDATE tableDeclarationPonderee " +
            "SET TD = TD * 10.0/8.0 "+
            "WHERE TD IS NOT NULL;";
        dbDeclaration.updateSQL(sql);
        // Pondération des TP
        sql="UPDATE tableDeclarationPonderee " +
            "SET TP = TP * 24.0/16.0 "+
            "WHERE (TP IS NOT NULL) "
            + "AND ((SELECT CoefTP FROM tableCoefTP WHERE tableCoefTP.EC_Code = tableDeclarationPonderee.Ects)='24.0');";
        dbDeclaration.updateSQL(sql);
        sql="UPDATE tableDeclarationPonderee " +
            "SET TP = TP * 28.5/16.0 "+
            "WHERE (TP IS NOT NULL) "
            + "AND ((SELECT CoefTP FROM tableCoefTP WHERE tableCoefTP.EC_Code = tableDeclarationPonderee.Ects)='28.5');";
        dbDeclaration.updateSQL(sql);

        // Sauvegarde de la table dans un fichier EXCEL
        sql = "SELECT NumIndividu, Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationPonderee " +
              "WHERE ((SELECT Statut FROM tableEnseignant WHERE tableEnseignant.Numero = tableDeclarationPonderee.NumIndividu)='PERM') "+  
              "ORDER BY Nom, Prenom, Ects;";
        res =  dbDeclaration.querySQL(sql);
        dbDeclaration.appendResultSetAsSheetToExcelFile(res, "Déclaration Permanents", path+"tableDeclaration.xlsx");
    //    sEFM = new sExcelFileManager();
    //    sEFM.newSheetFromResultSet("Déclaration Permanents", res);
    //    sEFM.saveAs(path+"tableDeclarationPondereePerm.xlsx");  
        
        // Sauvegarde de la table dans un fichier EXCEL
        sql = "SELECT NumIndividu, Ects, CM, TD, TP, Projet, Nom, Prenom "+
              "FROM tableDeclarationPonderee " +
              "WHERE ((SELECT Statut FROM tableEnseignant WHERE tableEnseignant.Numero = tableDeclarationPonderee.NumIndividu)='VACA') "+  
              "ORDER BY Nom, Prenom, Ects;";
        res =  dbDeclaration.querySQL(sql);
        dbDeclaration.appendResultSetAsSheetToExcelFile(res, "Déclaration Vacataires", path+"tableDeclaration.xlsx");
//        sEFM = new sExcelFileManager();
//        sEFM.newSheetFromResultSet("Déclaration Vacataires", res);
//        sEFM.saveAs(path+"tableDeclarationPondereeVaca.xlsx");  
        
        
        addMessage(2,"FIN OK.");
    }

    @FXML
    private void OnAction_BTN_3_XLS_SCOL(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            FXID_TXF_3_XLS_SCOL.setText(selectedFile.getPath());
        }
    }

    @FXML
    private void OnAction_BTN_3_GO(ActionEvent event) throws ClassNotFoundException, SQLException, IOException, InvalidFormatException, sExcelFileManagerException {
        
        addMessage(3,"Conversion d'index de colonne : 1 donne "+intToLetter(1)+" (Bon si A)\n");
        addMessage(3,"Conversion d'index de colonne : 26 donne "+intToLetter(26)+" (Bon si Z)\n");
        addMessage(3,"Conversion d'index de colonne : 27 donne "+intToLetter(27)+" (Bon si AA)\n");
        addMessage(3,"Conversion d'index de colonne : A donne "+letterToInt("A")+" (Bon si 1)\n");
        addMessage(3,"Conversion d'index de colonne : Z donne "+letterToInt("Z")+" (Bon si 26)\n");
        addMessage(3,"Conversion d'index de colonne : AA donne "+letterToInt("AA")+" (Bon si 27)\n");


        
        addMessage(3,"CONSTRUCTION DES TABLEAUX DE JURYS ... \n");
        
    /********************** 
              TRAITEMENTS
    ***********************/ 
        addMessage(3,"------- Début du traitement -------------\n");
        
        String path = "/Users/stephane/Documents/DRIVE.GOOGLE.MyDrive/DEVELOP/Java/NetBeansProjects/Excel-TestFiles/";

    // Création de la base
        addMessage(3,"Création de la table 'NotesEtudiants' et remplissage ..."); 
        sSQLiteManager dbDeclaration = new sSQLiteManager("sqlite_Jury.db");
        dbDeclaration.convertExcelFileIntoTable(path + "OR_Notes_EC_examens_Echanges_v1.xlsx", "NotesEtudiants");
        addMessage(3,"OK\n"); 

        
        addMessage(3,"Création du fichier 'testNotesBrutes.xlsx' ..."); 
        String sql = "SELECT DISTINCT Numero, Nom, Prenom, Code_ec, Note, Grade, Libelle_examen, Semestre "+
                        "FROM NotesEtudiants "+
                        "WHERE Semestre = 1 "+
                        "ORDER BY Nom, Prenom, Numero";
        ResultSet res =  dbDeclaration.querySQL(sql);
//        Workbook workbook = dbDeclaration.appendResultSetAsSheetToExcelFile(res, "Declaration Brute", path+"testDeclarationBrute.xlsx");
//        dbDeclaration.saveWorkbookToExcelFile(workbook, path+"testDeclarationBrute.xlsx");
        sExcelFileManager sEFM = new sExcelFileManager();
        sEFM.newSheetFromResultSet("Notes Brutes", res);
//        sEFM.setWorkbook(workbook);
        sEFM.saveAs(path+"testNotesBrutes.xlsx");
        addMessage(3,"OK\n"); 
        
        // Boucle sur toutes les lignes du resultat
        String precNUMERO="XXXX";
        while(res.next()) {
            String NUMERO=res.getString("Numero");
            if (NUMERO.equalsIgnoreCase(precNUMERO)) {
                
            }
            else {
                sql="UPDATE tableExcelTypeNotes " +
                    "SET NumeroTP = TP * 28.5/16.0 "+
            "WHERE (TP IS NOT NULL) "
            + "AND ((SELECT CoefTP FROM tableCoefTP WHERE tableCoefTP.EC_Code = tableDeclarationPonderee.Ects)='28.5');";
        dbDeclaration.updateSQL(sql);
                
            }
            String CODE_EC=res.getString("Code_ec");
        }
        

//        SQLiteStatement stmt = db.compileStatement("UPDATE users SET field1=?, field2=?...");
//        stmt.bindString(1, "value1");
//        stmt.bindString(2, "value2");
//        stmt.execute();

    }
}
