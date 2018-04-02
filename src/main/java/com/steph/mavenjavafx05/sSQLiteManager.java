/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.steph.mavenjavafx05;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.EmptyFileException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author stephane
 */
public class sSQLiteManager {
    
    private Connection m_Connection;
    private List<String> m_listTables;
    
    public sSQLiteManager() {
        
    }
    
    /**
     * Création de la base de données dans le fichier dbName
     * 
     * @param dbName 
     * @throws java.lang.ClassNotFoundException 
     * @throws java.sql.SQLException 
     */
    public sSQLiteManager(String dbName) throws ClassNotFoundException, SQLException {
        System.out.print("Création de la base "+dbName+" ... ");
        String driver = "org.sqlite.JDBC";
        Class.forName(driver);
        String dbUrl = "jdbc:sqlite:" + dbName;
        m_Connection = DriverManager.getConnection(dbUrl);
        m_listTables = new ArrayList<>();
        System.out.println("OK");      
    }

    /**
     * Renvoie le code character d'une valeur int :
     * 1 donne A, 2 donne B, ... 26 donne Z, 27 donne AA ...
     * @param i
     * @return 
     */
    public static String intToLetter(int i) {
        i=i-1;
        char charA = 'A';
        int ii=i;
        String s="";
        
        while (ii>=26) {
            ii=i;            
            i=ii%26;
            ii=ii-26;
            s=s+(char)(i+(int)charA);
        }
        i=ii;
        s=s+(char)(i+(int)charA);
        return s;
    } 
    
    /**
     * Renvoie la valeur int correspondant au code character  :
     * A donne 1, B donne 2, ... Z donne 26, AA donne 27 ...
     * @param i
     * @return 
     */
    public static int letterToInt(String s) {
        char charA = 'A';
        int i=0;
        
        for(int n=0; n<s.length(); n++) {
            i=i*26+(int)(s.charAt(n))-(int)charA+1;
        }
        return i;
    } 
        
    
    
    
    /**
     * Ajoute le nom de la table à la liste des tables de la base
     * 
     * @param table 
     */
    private void addTableToTableList(String table) {
        m_listTables.add(table);
    }

    /**
     * Création avec écrasement eventuel SANS remplissage d'une table 'table' 
     * avec les noms de colonnes contenus dans listColumn  
     * @param table
     * @param listColumn
     * @throws SQLException 
     */
    public void createTableDropIfExists(String table, List<String> listColumn) throws SQLException {
        System.out.print("Création de la table "+table+" ... ");
        String sql;
        Statement stQuery = m_Connection.createStatement();
        Statement stUpdate = m_Connection.createStatement();
        stUpdate.executeUpdate("DROP TABLE IF EXISTS "+ table);
        sql = "CREATE TABLE "+table+" (";
//        sql = sql + "Id INTEGER PRIMARY KEY AUTOINCREMENT, ";
        for (String s : listColumn) {
            sql = sql + s + ", ";
        }
        sql=sql.substring(0, sql.length()-2);
        sql = sql+")";
        stUpdate.executeUpdate(sql);
        addTableToTableList(table);
        System.out.println("OK");
    }

    /**
     * Ajoute une ligne à la table 'table' en mettant, dans les colonnes listColumn, 
     * les valeurs de 'listValues' dans l'ordre correspondant
     * @param table
     * @param listColumn
     * @param listValues
     * @throws SQLException 
     */
    void appendRowIntoTable(String table, List<String> listColumn, List<String> listValues) throws SQLException {
        //System.out.print("INSERT INTO " + table + " ... ");
        String sql;
        //Statement stQuery = m_Connection.createStatement();
        Statement stUpdate = m_Connection.createStatement();
        sql = "INSERT INTO  "+table+" (";
        for (String column : listColumn) {
            sql = sql + column + ", ";
        }
        sql=sql.substring(0, sql.length()-2);
        sql = sql+") VALUES (";
        for (String value : listValues) {
            sql = sql + "'" + value + "', ";
        }
        sql=sql.substring(0, sql.length()-2);
        sql = sql+");";
        //System.out.println("sql = "+sql);
        stUpdate.executeUpdate(sql);
    }
    
    
    /**
     * Execute une requete SQL de type query : SELECT ...
     * @param query
     * @return
     * @throws SQLException 
     */
    public ResultSet querySQL(String query) throws SQLException {
        Statement stQuery = m_Connection.createStatement(); 
        ResultSet rs = stQuery.executeQuery(query);
        return rs;
    }

    /**
     * Execute une requete SQL de type update : UPDATE, CREATE, ...
     * @param sql
     * @throws SQLException 
     */
    public void updateSQL(String sql) throws SQLException {
        Statement stUpdate = m_Connection.createStatement(); 
        stUpdate.executeUpdate(sql);
    }
    
    
    /**
     * Création (avec écrasement éventuel si existe ...) ET remplissage 
     * d'une table dans la base à partir d'un fichier Excel et de la feuille 
     * sheetnumber = 0 en reprenant directement les noms de colonnes du fichier Excel 
     * pour les noms de colonnes de la table. 
     * Suppose un fichier Excel avec En Tete, donc la ligne n°1 n'est pas lu.
     * @param table
     * @param excelfile 
     * @throws java.io.IOException 
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException 
     * @throws com.steph.mavenjavafx05.sExcelFileManagerException 
     * @throws java.sql.SQLException 
     */
    public void convertExcelFileIntoTable(String excelfile, String table) throws IOException, InvalidFormatException, sExcelFileManagerException, SQLException {
        sExcelFileManager sEFM = new sExcelFileManager(excelfile);
        List<String> listColumn = sEFM.getSheetHeader(0);
        convertExcelFileIntoTable(excelfile, table, 0, listColumn);
    }

    /**
     * Création (avec écrasement éventuel si existe ...) ET remplissage 
     * d'une table dans la base à partir d'un fichier Excel et de la feuille 
     * sheetnumber = 0 en remplaçant les noms de colonnes du fichier Excel 
     * par les noms de colonnes de listColumn dans la table
     * Suppose un fichier Excel avec En Tete, donc la ligne n°1 n'est pas lu.
     * @param table
     * @param excelfile 
     * @param listColumn 
     * @throws java.io.IOException 
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException 
     * @throws com.steph.mavenjavafx05.sExcelFileManagerException 
     * @throws java.sql.SQLException 
     */
    public void convertExcelFileIntoTable(String excelfile, String table, List<String> listColumn) throws IOException, InvalidFormatException, sExcelFileManagerException, SQLException {
        sExcelFileManager sEFM = new sExcelFileManager(excelfile);
        convertExcelFileIntoTable(excelfile, table, 0, listColumn);
    }
  
    /**
     * Création (avec écrasement éventuel si existe ...) ET remplissage 
     * d'une table dans la base à partir d'un fichier Excel et de la feuille 
     * sheetnumber en reprenant directement les noms de colonnes du fichier Excel 
     * pour les noms de colonnes de la table
     * Suppose un fichier Excel avec En Tete, donc la ligne n°1 n'est pas lu.
     * @param table
     * @param excelfile 
     * @param sheetnumber 
     * @throws java.io.IOException 
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException 
     * @throws com.steph.mavenjavafx05.sExcelFileManagerException 
     * @throws java.sql.SQLException 
     */
    public void convertExcelFileIntoTable(String excelfile, String table, int sheetnumber) throws IOException, InvalidFormatException, sExcelFileManagerException, SQLException {
        sExcelFileManager sEFM = new sExcelFileManager(excelfile);
        List<String> listColumn = sEFM.getSheetHeader(sheetnumber);
        List<String> listNewColumn = new ArrayList<>();
        listNewColumn.add("Id INTEGER PRIMARY KEY AUTOINCREMENT");
        for (String column : listColumn) {
            listNewColumn.add(column + " CHAR(255)");            
        }
        convertExcelFileIntoTable(excelfile, table, sheetnumber, listNewColumn);
    }

    /**
     * Création (avec écrasement éventuel si existe ...) ET remplissage 
     * d'une table dans la base à partir d'un fichier Excel et de la feuille 
     * sheetnumber en remplaçant les noms de colonnes du fichier Excel 
     * par les noms de colonnes de la liste listColumn dans la table
     * Suppose un fichier Excel avec En Tete, donc la ligne n°1 n'est pas lu.
     * @param table
     * @param excelfile 
     * @param sheetnumber 
     * @param listColumn 
     * @throws java.sql.SQLException 
     * @throws java.io.IOException 
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException 
     */
    public void convertExcelFileIntoTable(String excelfile, String table, int sheetnumber, List<String> listColumn) throws SQLException, IOException, InvalidFormatException, sExcelFileManagerException {
        System.out.print("Creation de la table "+table+" ... ");
        createTableDropIfExists(table, listColumn);

        sExcelFileManager sEFM = new sExcelFileManager(excelfile);
        sEFM.setCurrentSheetFromSheetNumber(sheetnumber);
        // Boucle sur les lignes en sautant la première (Header)
        for (int ligne=1; ligne<sEFM.getCurrentSheetSize(); ligne++) {
            List<String> listValues = new ArrayList<>();
            for (int colonne=0; colonne<listColumn.size(); colonne++) {
                //System.out.println(ligne + " "+ colonne);
                String value=sEFM.getDataFromCurrentSheet(ligne, colonne);
                // Remplacement de ' par $ pour les requetes SQL
                String newvalue = value.replace("'", "$"); 
                listValues.add(newvalue);
            }
            appendRowIntoTable(table, listColumn, listValues);            
        }
        System.out.println("OK");
        
    }

    
    
    
    
    
    /**
     * Création ET remplissage de la table 'table' à partir du fichier EXCEL excelfile
     * et la feuille de nom sheetname ou sheet numéro 0 si 'sheetname' n'existe pas
     * 
     * @param table
     * @param excelfile
     * @param sheetname
     * @throws IOException
     * @throws InvalidFormatException
     * @throws SQLException 
     */
    public void createAndFillTableFromExcelSheet(String table, String excelfile, String sheetname) throws IOException, InvalidFormatException, SQLException { 
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(excelfile));
        // Getting the Sheet at index sheetnumber :first is zero
        int i=0;
        int index=0;
        for (Sheet sheet : workbook) {
            if (sheet.getSheetName().equalsIgnoreCase(sheetname)) {
                index=i;
                break;
            }
            i++;
        }
        // Closing the workbook
        workbook.close();
        System.out.println("Table : " + table + " à l'index : " + index);
        createAndFillTableFromExcelSheet(table, excelfile, index);
    }
    
    
    /**
     * Création ET remplissage de la table 'table' à partir du fichier EXCEL excelfile
     * et la feuille sheetnumber
     * 
     * @param table
     * @param excelfile
     * @param sheetnumber 
     * @throws java.io.IOException 
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException 
     * @throws java.sql.SQLException 
     */
    public void createAndFillTableFromExcelSheet(String table, String excelfile, int sheetnumber) throws IOException, InvalidFormatException, SQLException {
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(excelfile));
        Sheet sheet = workbook.getSheetAt(sheetnumber);
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();        
        
        // Getting the Sheet at index sheetnumber :first is zero
        int numberOfRows=sheet.getLastRowNum();
        int numberOfColumn = (int)(sheet.getRow(0).getLastCellNum()); // Renvoie la dernière colonne + 1
        // On peut comparer avec getPhysicalNumberOfCells() qui donne le nombre de colonnes non nulles dans la Row

        // Récupération du header
        List<String> listColumn = new ArrayList<>();
        for (int col=0; col<numberOfColumn; col++) {
            Cell cell = sheet.getRow(0).getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//RETURN_BLANK_AS_NULL))
            String cellValue = dataFormatter.formatCellValue(cell);
            listColumn.add(cellValue);
        }
        System.out.print("Creation de la table "+table+" ... ");
        createTableDropIfExists(table, listColumn);
        System.out.println("OK");

        System.out.print("Remplissage de la table "+table+" ... ");
        // Boucle sur les lignes en sautant la première (Header)
        for (int ligne=1; ligne<numberOfRows; ligne++) {
            List<String> listValues = new ArrayList<>();
            for (int col=0; col<numberOfColumn; col++) {
                Cell cell = sheet.getRow(col).getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//RETURN_BLANK_AS_NULL))
                String cellValue = dataFormatter.formatCellValue(cell);
                // Remplacement de ' par $ pour inclure la valeur dans les requetes SQL
                String newvalue = cellValue.replace("'", "$"); 
                listValues.add(newvalue);
            }
            appendRowIntoTable(table, listColumn, listValues);            
        }
        // Closing the workbook
        workbook.close();
        System.out.println("OK");
    }
    
    /**
     * Ajout des lignes à la table existante 'table'. Les headers de colonnes du fichier doivent correspondre aux
     * noms de colonnes de la table
     * @param table
     * @param excelfile
     * @param sheetname
     * @throws IOException
     * @throws InvalidFormatException
     * @throws SQLException 
     */
    public void appendToTableFromExcelSheet(String table, String excelfile, String sheetname) throws IOException, InvalidFormatException, SQLException { 
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(excelfile));
        // Getting the Sheet at index sheetnumber :first is zero
        int i=0;
        int index=0;
        for (Sheet sheet : workbook) {
            if (sheet.getSheetName().equalsIgnoreCase(sheetname)) {
                index=i;
                break;
            }
            i++;
        }
        // Closing the workbook
        workbook.close();
        appendToTableFromExcelSheet(table, excelfile, index);
    }

    public void appendToTableFromExcelSheet(String table, String excelfile, int sheetnumber) throws IOException, InvalidFormatException, SQLException {
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(excelfile));
        Sheet sheet = workbook.getSheetAt(sheetnumber);
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();        
        
        // Getting the Sheet at index sheetnumber :first is zero
        int numberOfRows=sheet.getLastRowNum();
        int numberOfColumn = (int)(sheet.getRow(0).getLastCellNum()); // Renvoie la dernière colonne + 1
        // On peut comparer avec getPhysicalNumberOfCells() qui donne le nombre de colonnes non nulles dans la Row

        // Récupération du header
        List<String> listColumn = new ArrayList<>();
        for (int col=0; col<numberOfColumn; col++) {
            Cell cell = sheet.getRow(0).getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//RETURN_BLANK_AS_NULL))
            String cellValue = dataFormatter.formatCellValue(cell);
            listColumn.add(cellValue);
        }
        System.out.print("Creation de la table "+table+" ... ");
        createTableDropIfExists(table, listColumn);
        System.out.println("OK");

        System.out.print("Remplissage de la table "+table+" ... ");
        // Boucle sur les lignes en sautant la première (Header)
        for (int ligne=1; ligne<numberOfRows; ligne++) {
            List<String> listValues = new ArrayList<>();
            for (int col=0; col<numberOfColumn; col++) {
                Cell cell = sheet.getRow(col).getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//RETURN_BLANK_AS_NULL))
                String cellValue = dataFormatter.formatCellValue(cell);
                // Remplacement de ' par $ pour inclure la valeur dans les requetes SQL
                String newvalue = cellValue.replace("'", "$"); 
                listValues.add(newvalue);
            }
            appendRowIntoTable(table, listColumn, listValues);            
        }
        // Closing the workbook
        workbook.close();
        System.out.println("OK");
    }
    
    
    // Partie pour le fichier de sortie en EXCEL
    
    private Workbook getWorkbook(String excelfile) {
        Workbook workbook=null;        
        if (new File(excelfile).isFile()) {
            // Le fichier existe
            // Creating a Workbook from an Excel file (.xls or .xlsx)
            try {
                workbook = WorkbookFactory.create(new File(excelfile));
            }
            catch (IOException e) {
                System.out.println("CATCH IOException: " +e);
                workbook=null;
            }
            catch (InvalidFormatException e) {
                System.out.println("CATCH InvalidFormatException: " +e);  
                workbook=null;
            }
            catch (EmptyFileException e) {
                System.out.println("CATCH EmptyFileException: " +e);   
                workbook=null;
            }
        } 
        
        if (workbook==null) {
            // Le workbook n'existe pas encore 
            if (excelfile.toUpperCase().endsWith("XLSX")) {
                workbook = new XSSFWorkbook();
            } else if (excelfile.toUpperCase().endsWith("XLS")) {
                workbook = new HSSFWorkbook();
            } else {
                throw new IllegalArgumentException("The specified file is not Excel file");
            }
        }
        return workbook;
    }

    
    /**
     * Ajoute le contenu du ResultSet 'res' en tant que la feuille Excel 'sheetname'
     * dans le fichier 'excelfile'
     * @param res
     * @param sheetname
     * @param excelfile
     */
    public Workbook appendResultSetAsSheetToExcelFile(ResultSet res, String sheetname, String excelfile) throws IOException, InvalidFormatException, SQLException {
        
        Workbook workbook = getWorkbook(excelfile);
        Sheet sheet = workbook.getSheet(sheetname);
        if(sheet != null)   {
            int index = workbook.getSheetIndex(sheet);
            workbook.removeSheetAt(index);
        }        
        sheet=workbook.createSheet(sheetname);

        // Exploitation des ResultSet
//        String sql  = "SELECT * FROM MATABLE"; 
//        Statement statement = connection.createStatement(); 
//        ResultSet resultat = statement.executeQuery(sql); 
//        ResultSetMetaData metadata = resultat.getMetaData(); 
//        int nombreColonnes = metadata.getColumnCount(); 
//        System.out.println("Ce ResultSet contient "+nombreColonnes+" colonnes.");        
        
        ResultSetMetaData metadata = res.getMetaData();   // Données complémentaires pour le ResultSet 
        int numberOfColumns = metadata.getColumnCount(); // Nombre de colonnes dans le résultat
        List<String> listHeaderName = new ArrayList<>(numberOfColumns); 
        for(int i = 1; i <= numberOfColumns; i++){ 
            listHeaderName.add(metadata.getColumnName(i)); // index des ResultSet commence à 1
        } 

        // Enregistrement du Header des colommes
        Row rowHeader = sheet.createRow(0);
        int index=0;
        for (String headername : listHeaderName) {
            Cell cell = rowHeader.createCell(index);
            cell.setCellValue(headername);
            index++;
        }
        int rowCount = 1; 
        while (res.next()) {
            Row row = sheet.createRow(rowCount);
            for (int i=0; i<numberOfColumns; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(res.getString(i+1));
            }
            rowCount++;
        }
        
        for(Sheet sh : workbook) {
            System.out.println(sh.getSheetName());
        }

        return workbook;

    }

    public void saveWorkbookToExcelFile(Workbook workbook, String excelfile) {
        try {
            System.out.println("FileOutputStream outputStream = new FileOutputStream(excelfile);");
            FileOutputStream outputStream = new FileOutputStream(excelfile);
            System.out.println("workbook.write(outputStream);");
            workbook.write(outputStream);
            System.out.println("workbook.close();");
//            workbook.close();
            System.out.println("OK");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    
    
   /****************************
    * 
    * Methodes non génériques  
    * 
    ***************************/ 
    
    
    
    void addTableYann(String tableYann, String fileName, String annee) throws IOException, InvalidFormatException, sExcelFileManagerException, SQLException {
        List<String> listColumn = new ArrayList<>();

        System.out.print("Ajout de la table des données de Yann ... ");
        sExcelFileManager sEFM_tableCreneaux = new sExcelFileManager(fileName);
        sEFM_tableCreneaux.setCurrentSheetFromSheetNumber(0);
        listColumn.add("Jour");
        listColumn.add("Module");
        listColumn.add("Type");
        listColumn.add("Enseignant");
        listColumn.add("Groupe");
        listColumn.add("Heure_At_Salle");
        listColumn.add("Heure_Debut");
        listColumn.add("Heure_Fin");
        listColumn.add("Salle");
        listColumn.add("DureeMinutes");
        listColumn.add("Annee");

        for (int ligne=1; ligne<sEFM_tableCreneaux.getCurrentSheetSize(); ligne++) {
            List<String> listValues = new ArrayList<>();
            listValues.add(sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Jour").toUpperCase());
            String module = sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Module");
            String newmodule = module.replace("'", "$");            
            listValues.add(newmodule);
            listValues.add(sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Type").toUpperCase());
            listValues.add(sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Enseignant").toUpperCase());
            listValues.add(sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Groupe").toUpperCase());
            listValues.add(sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Heure @ Salle").toLowerCase());
            
            // Décomposition de 08h00~10h00 @ 503:Gaston Berger
            String Heure_At_Salle = sEFM_tableCreneaux.getDataFromCurrentSheet(ligne, "Heure @ Salle").toLowerCase();
            String[] result1 = Heure_At_Salle.split("~");
            String Heure_Debut = result1[0].trim();
            String[] result2 = result1[1].split("@");
            String Heure_Fin = result2[0].trim();
            String Salle = "";
            if (result2.length>1) {
                Salle=result2[1].trim();
            }
             
            
            // Calcul de la Durée en minutes
            String[] resultDebut = Heure_Debut.split("h");
            String[] resultFin =   Heure_Fin.split("h");
            int duree = Integer.parseInt(resultFin[0])*60 + Integer.parseInt(resultFin[1]) - 
                    Integer.parseInt(resultDebut[0])*60 - Integer.parseInt(resultDebut[1]);
            
            listValues.add(Heure_Debut);
            listValues.add(Heure_Fin);
            listValues.add(Salle);
            listValues.add(String.valueOf(duree));
            listValues.add(annee);
            
//            for (String s : listValues) {
//                System.out.println(s);
//            }
//            System.out.println("--------");
            
            appendRowIntoTable(tableYann, listColumn, listValues);            
        }
        System.out.println("OK");
    }

    void createTableYannDropIfExists(String table) throws SQLException {
    //create table 
        List<String> listColumn = new ArrayList<>();
        listColumn.add("Id INTEGER PRIMARY KEY AUTOINCREMENT");
        listColumn.add("Jour           CHAR(16)   NOT NULL");
        listColumn.add("Module         CHAR(255)  NOT NULL");
        listColumn.add("Type           CHAR(8)    NOT NULL");
        listColumn.add("Enseignant     CHAR(64)");
        listColumn.add("Groupe         CHAR(64)");
        listColumn.add("Heure_At_Salle CHAR(64)");
        listColumn.add("Heure_Debut    CHAR(8)");
        listColumn.add("Heure_Fin      CHAR(8)");
        listColumn.add("DureeMinutes   CHAR(8)");
        listColumn.add("Salle          CHAR(64)");
        listColumn.add("Annee          CHAR(8)");
        listColumn.add("Numero         CHAR(32)");
        listColumn.add("Nom            CHAR(64)");
        listColumn.add("Prenom         CHAR(64)");
        listColumn.add("EC_Code        CHAR(64)");
        createTableDropIfExists(table, listColumn);
    }
    
    
    
    void createTableMaquetteDropIfExists(String tableMaquette) throws SQLException {
    //create table 
        List<String> listSQL = new ArrayList<>();
        listSQL.add("Id INTEGER PRIMARY KEY AUTOINCREMENT");
        listSQL.add("EC_Code        CHAR(16)   NOT NULL");
        listSQL.add("Module         CHAR(255)  NOT NULL");
        listSQL.add("Module_Complet CHAR(255)  NOT NULL");
        listSQL.add("Heure          CHAR(8)");
        listSQL.add("ECTS           CHAR(8)");
        createTableDropIfExists(tableMaquette, listSQL);
    }

    void addTableMaquette(String tableMaquette, String fileName) throws IOException, InvalidFormatException, sExcelFileManagerException, SQLException {
        List<String> listColumn = new ArrayList<>();

        System.out.print("INSERT INTO tableMaquette ... ");
        sExcelFileManager sEFM_tableMaquette = new sExcelFileManager(fileName);
        sEFM_tableMaquette.setCurrentSheetFromSheetNumber(0);
        listColumn.add("EC_Code");
        listColumn.add("Module");
        listColumn.add("Module_Complet");
        listColumn.add("Heure");
        listColumn.add("ECTS");

        for (int ligne=1; ligne<sEFM_tableMaquette.getCurrentSheetSize(); ligne++) {
            List<String> listValues = new ArrayList<>();
            listValues.add(sEFM_tableMaquette.getDataFromCurrentSheet(ligne, "EC_Code").toUpperCase());
            String module = sEFM_tableMaquette.getDataFromCurrentSheet(ligne, "Module");
            String newmodule = module.replace("'", "$");            
            listValues.add(newmodule);
            
            module = sEFM_tableMaquette.getDataFromCurrentSheet(ligne, "Module_Complet");
            newmodule = module.replace("'", "$");            
            listValues.add(newmodule);
            
            listValues.add(sEFM_tableMaquette.getDataFromCurrentSheet(ligne, "Heures").toUpperCase());
            listValues.add(sEFM_tableMaquette.getDataFromCurrentSheet(ligne, "ECTS").toUpperCase());
            
            for (String s : listValues) {
                System.out.println(s);
            }
            System.out.println("--------");
            
            appendRowIntoTable(tableMaquette, listColumn, listValues);            
        }
        System.out.println("OK");
    }

    void complementTableCreneauxYann(String tableCreneaux, String tableMaquette) throws SQLException {
        String SQL; // = "UPDATE tableCreneaux SET tableCreneaux.EC_Code = tableMaquette.EC_Code FROM tableCreneaux  INNER JOIN  tableMaquette ON tableCreneaux.Module = tableMaquette.Module";
        SQL="UPDATE tableCreneaux " +
            "SET EC_Code = (SELECT EC_Code " +
            "               FROM tableMaquette " +
            "               WHERE UPPER(tableCreneaux.Module) = UPPER(tableMaquette.Module)); ";
        Statement stUpdate = m_Connection.createStatement(); 
        stUpdate.executeUpdate(SQL);
        
        SQL="UPDATE tableCreneaux " +
            "SET Nom = (SELECT D.Nom "+
            "           FROM tableEnseignant D " +
            "           WHERE UPPER(tableCreneaux.Enseignant) = UPPER(D.Enseignant)), "+
            "    Prenom = (SELECT D.Prenom "+
            "              FROM tableEnseignant D " +
            "              WHERE UPPER(tableCreneaux.Enseignant) = UPPER(D.Enseignant)), "+
            "    Numero = (SELECT D.Numero "+
            "              FROM tableEnseignant D " +
            "              WHERE UPPER(tableCreneaux.Enseignant) = UPPER(D.Enseignant)); ";
//            WHERE EXISTS (SELECT D.Enseignant 
//                 FROM Table_2 d
//                 WHERE Table_1.Id = d.Id
//                   AND Table_1.EmailId <> d.EmailId
//                   AND d.EmailId IS NOT NULL
//                 );

//        SQL="UPDATE tableCreneaux " +
//            "SET tableCreneaux.Nom = tableEnseignant.Nom " +
//            "    tableCreneaux.Prenom = tableEnseignant.Prenom " +
//            "    tableCreneaux.Numero = tableEnseignant.Numero " +
//            "WHERE UPPER(tableCreneaux.Enseignant) = UPPER(tableEnseignant.Enseignant); ";
        stUpdate = m_Connection.createStatement(); 
        stUpdate.executeUpdate(SQL);
    }

    void saveWorkbookToExcelFile(Workbook workbook) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }


}
