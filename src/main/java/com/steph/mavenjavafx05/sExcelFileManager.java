/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.steph.mavenjavafx05;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



/**
 *
 * @author stephane
 * Inspiré de :
 * https://www.callicoder.com/java-read-excel-file-apache-poi/
 */

public class sExcelFileManager {
    

    private String                      m_ExcelReadFileName;
    private String                      m_ExcelSaveFileName;
    private Workbook                    m_Workbook=null;
    private String                      m_ExcelFileType;
    private List<String>                m_ListSheetsName;
    private List<List<List<String>>>    m_Data = new ArrayList<>();
    private int                         m_CurrentSheetNumber = 0;

    // Constructeur par défaut
    public sExcelFileManager() {
        m_ExcelReadFileName="";
        
        m_ExcelSaveFileName="";
        m_ExcelFileType = "XLSX";
    }
    
    public void setExcelFileType(String type) {
        if (m_Workbook==null) {
            if (type.toUpperCase().endsWith("XLSX")) {
                m_ExcelFileType = "XLSX";
            }
            if (type.toUpperCase().endsWith("XLS")) {
                m_ExcelFileType = "XLS";
            }
        }
    }
    
    public void setSaveFileName(String filename) throws IOException {
        if (m_Workbook==null) {
            m_ExcelSaveFileName=filename; 
            setExcelFileType(filename);
        }
        else {
            if (!filename.toUpperCase().endsWith(m_ExcelFileType)) { 
                if (filename.contains(".")) {
                    filename = filename.substring(0, filename.lastIndexOf('.'));
                }
                filename += "."+m_ExcelFileType;
            }
            m_ExcelSaveFileName=filename;
        }
    }
    
    /**
     *
     * @param excelfilename
     * @throws IOException
     * @throws InvalidFormatException
     */
    public sExcelFileManager(String excelfilename) throws IOException, InvalidFormatException {
        m_ExcelReadFileName = excelfilename;
        
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        m_Workbook = WorkbookFactory.create(new File(m_ExcelReadFileName));
        m_ListSheetsName = new ArrayList<String>();
        
        
        /*****************************************************************
           Iterating over all the sheets in the workbook (Multiple ways)
         *****************************************************************/
        
        // 1. You can obtain a sheetIterator and iterate over it
        //Iterator<Sheet> sheetIterator = m_Workbook.sheetIterator();
        //System.out.println("Retrieving Sheets using Iterator");
        //while (sheetIterator.hasNext()) {
        //    Sheet sheet = sheetIterator.next();
        //    m_ListSheetsName.add(sheet.getSheetName());
        //}

        // 2. Or you can use a for-each loop
        //System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: m_Workbook) {
            m_ListSheetsName.add(sheet.getSheetName());
        }

        // 3. Or you can use a Java 8 forEach with lambda
        //System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        //m_Workbook.forEach(sheet -> {
        //  System.out.println("=> " + sheet.getSheetName());
        //    m_ListSheetsName.add(sheet.getSheetName());
        //});                

        
        
        /********************************************************************
           Iterating over all the rows and columns in a Sheet (Multiple ways)
         ********************************************************************/
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // Getting the Sheet at index sheetnumber :first is zero
        for (Sheet sheet : m_Workbook) {
            //Sheet sheet = m_Workbook.getSheetAt(sheetnumber);
            int numberOfColumnInCurrentSheet; // Renvoie la dernière colonne + 1
            // On peut comparer avec getPhysicalNumberOfCells() qui donne le nombre de colonnes non nulles dans la Row
            numberOfColumnInCurrentSheet = (int)(sheet.getRow(0).getLastCellNum());
            
        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        //System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        //Iterator<Row> rowIterator = sheet.rowIterator();
        //while (rowIterator.hasNext()) {
        //    Row row = rowIterator.next();
        //
        //    // Now let's iterate over the columns of the current row
        //    Iterator<Cell> cellIterator = row.cellIterator();
        //
        //    while (cellIterator.hasNext()) {
        //        Cell cell = cellIterator.next();
        //        String cellValue = dataFormatter.formatCellValue(cell);
        //        System.out.print(cellValue + "\t");
        //    }
        //    System.out.println();
        //}

            // 2. Or you can use a for-each loop to iterate over the rows and columns
            //System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
            List<List<String>> sheetData = new ArrayList<>(); 
            
            for (Row row: sheet) {
                
                List<String> rowData = new ArrayList<>();
//                for(Cell cell: row) {
//                    String cellValue = dataFormatter.formatCellValue(cell);
//                    rowData.add(cellValue);
//                }
//                for(int cn=0; cn<row.getLastCellNum(); cn++) {
                for(int cn=0; cn<numberOfColumnInCurrentSheet; cn++) {
                    // If the cell is missing from the file, generate a blank one
                    // (Works by specifying a MissingCellPolicy)
                    Cell cell = row.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellValue = dataFormatter.formatCellValue(cell);
                    rowData.add(cellValue);  
                }
                sheetData.add(rowData);
            }
            m_Data.add(sheetData);
            
        // 3. Or you can use Java 8 forEach loop with lambda
//        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
//        sheet.forEach(row -> {
//            row.forEach(cell -> {
//                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + "\t");
//            });
//            System.out.println();
//        });
        }

        // Closing the workbook
        m_Workbook.close();

    }
    
    /**
     *
     * @param excelfilename
     * @throws IOException
     * @throws InvalidFormatException
     */
    public List<List<String>> getSheetFromSheetName(String sheetname) throws sExcelFileManagerException {
        int sheetnumber=0;
        for (String s :  m_ListSheetsName ) {
            if (sheetname.equalsIgnoreCase(s)) {
                return m_Data.get(sheetnumber);
            }
            sheetnumber++;
        }
            // Envoi d'une exception SheetNotFound() : sheetname does not exist
            throw new sExcelFileManagerException("'sheetname' does not exist");
    }
    
    public List<List<String>> getSheetFromSheetNumber(int sheetnumber) throws sExcelFileManagerException {
        if (sheetnumber < m_ListSheetsName.size()) {
            return m_Data.get(sheetnumber);
        } 
        else {
            // Envoi d'une exception SheetNotFound() : sheetnumber out of range
            throw new sExcelFileManagerException("'sheetnumber' out of range in getSheetHeader");
        }
        
    }
        
    
    public int getNumberOfSheets() {
        return m_Workbook.getNumberOfSheets();
    }
    
    public List<String> getSheetsName() {
        return m_ListSheetsName;
    }

    /**
     *
     * @param sheetnumber
     * @return 
     * @throws com.steph.sexcelfilemanager.sExcelFileManagerException
     * @throws IOException
     * @throws InvalidFormatException
     */
    public List<String> getSheetHeader(int sheetnumber) throws sExcelFileManagerException {
        // test de la valeur de sheetnumber et envoi d'une exception
        if ((sheetnumber>=0) && (sheetnumber<m_ListSheetsName.size())) {
            return getSheetHeader(getSheetFromSheetNumber(sheetnumber));
        }
        else {
            throw new sExcelFileManagerException("'sheetnumber' out of range in getSheetHeader");
        }
    }
    public List<String> getSheetHeader(String sheetname) throws sExcelFileManagerException {
        // test de la valeur de sheetnumber et envoi d'une exception
        return getSheetHeader(getSheetFromSheetName(sheetname));
    }
    public List<String> getSheetHeader(List<List<String>> sheet) {
        return sheet.get(0);
    }
    
    public void setCurrentSheetFromSheetName(String sheetname) throws sExcelFileManagerException {
        int sheetnumber=0;
        for (String s :  m_ListSheetsName ) {
            if (sheetname.equalsIgnoreCase(s)) {
                m_CurrentSheetNumber = sheetnumber;
                return;
            }
            sheetnumber++;
        }
        // Envoi d'une exception SheetNotFound() : sheetname does not exist
        throw new sExcelFileManagerException("'sheetname' does not exist");
        
    }

    public void setCurrentSheetFromSheetNumber(int sheetnumber) throws sExcelFileManagerException {
        // test de la valeur de sheetnumber et envoi d'une exception
        if ((sheetnumber>=0) && (sheetnumber<m_ListSheetsName.size())) {
            m_CurrentSheetNumber = sheetnumber;
        }
        else {
            throw new sExcelFileManagerException("'sheetnumber' out of range in setCurrentSheetFromNumber");
        }
    }
    
    public int getColumnNumberFromHeader(String headername) throws sExcelFileManagerException {
        // Iterates the data and print it out to the console.
        int column = 0;
        for (String name : getSheetHeader(getSheetFromSheetNumber(m_CurrentSheetNumber))) {
            if (name.equalsIgnoreCase(headername)) {
                return column;
            }
            column++;
        }
        throw new sExcelFileManagerException("'headername' does not exist in getColumnNumberFromHeader()");
    }
    
    public List<String> getColumnFromHeader(String headername) throws sExcelFileManagerException {
        return getColumnFromNumber(getColumnNumberFromHeader(headername), true);
    }

    public List<String> getColumnFromNumber(int columnnumber, boolean withHeader) throws sExcelFileManagerException {
        // Iterates the data and print it out to the console.
        int start=1;
        if (withHeader) start=0;
        List<String> list = new ArrayList<>();
        for (List<String> data : getSheetFromSheetNumber(m_CurrentSheetNumber)) {
            list.add(data.get(columnnumber));
        }     
        return list;
    }
    
    public int getCurrentSheetSize() {
        return m_Data.get(m_CurrentSheetNumber).size();
    }
    
    public String getDataFromCurrentSheet(int ligne, String header) throws sExcelFileManagerException {
        int column = getColumnNumberFromHeader(header);
        return m_Data.get(m_CurrentSheetNumber).get(ligne).get(column);
    }
    public String getDataFromCurrentSheet(int ligne, int column) {
        return m_Data.get(m_CurrentSheetNumber).get(ligne).get(column);
    }
    
    public String getDataFromCurrentSheetFirstColumnValue(String firstColumnValue, String header) throws sExcelFileManagerException {
        int column = getColumnNumberFromHeader(header);
        for (List<String> list : m_Data.get(m_CurrentSheetNumber)) {
            if (list.get(0).equalsIgnoreCase(firstColumnValue)) {
                return list.get(column);
            }
        }
        return "";
    }
    
    private static void showSheetData(List<List<String>> sheet) {
        showSheetData(sheet, true); 
    }
    private static void showSheetData(List<List<String>> sheet, boolean withHeader) {
        // Iterates the data and print it out to the console.
        int start=1;
        if (withHeader) start=0;
        for (List<String> data : sheet) {
            for (int j = start; j < data.size(); j++) {
                String cell = data.get(j);
                System.out.print(cell + "\t");
            }
            System.out.println("");
        }
    }
    private static void showListData(List<String> list) {
        // Iterates the data and print it out to the console.
        for (String data : list) {
            System.out.print(data);
        }
    }
  
    



// Partie pour le fichier de sortie en EXCEL
    
    private Workbook getWorkbook() throws IOException {
        Workbook workbook = null;
        if (m_ExcelFileType.toUpperCase().endsWith("XLSX")) {
            workbook = new XSSFWorkbook();
        } else if (m_ExcelFileType.toUpperCase().endsWith("XLS")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
        return workbook;
    }

    /**
     *
     * @param sheettitle
     * @param rs
     * @throws IOException
     * @throws SQLException
     * 
     * Permet d'enregister le résultat d'une requète SQL dans une feuille EXCEL en conservant les 
     * entêtes de colonnes dans EXCEL identiques à ceux des colonnes de la base SQL 
     */
    public void newSheetFromResultSet(String sheettitle, ResultSet rs ) throws IOException, SQLException {
        
        // Exploitation des ResultSet
//        String sql  = "SELECT * FROM MATABLE"; 
//        Statement statement = connection.createStatement(); 
//        ResultSet resultat = statement.executeQuery(sql); 
//        ResultSetMetaData metadata = resultat.getMetaData(); 
//        int nombreColonnes = metadata.getColumnCount(); 
//        System.out.println("Ce ResultSet contient "+nombreColonnes+" colonnes.");        
        
        ResultSetMetaData metadata = rs.getMetaData();   // Données complémentaires pour le ResultSet 
        int numberOfColumns = metadata.getColumnCount(); // Nombre de colonnes dans le résultat
        List<String> listHeaderName = new ArrayList<>(numberOfColumns); 
        for(int i = 1; i <= numberOfColumns; i++){ 
            listHeaderName.add(metadata.getColumnName(i)); // index des ResultSet commence à 1
        } 

        if (m_Workbook==null) {
            m_Workbook=getWorkbook();
        }
        Sheet sheet = m_Workbook.createSheet(sheettitle); // On ajoute une feuille au fichier Excel

        // Enregistrement du Header des colommes
        Row rowHeader = sheet.createRow(0);
        int index=0;
        for (String headername : listHeaderName) {
            Cell cell = rowHeader.createCell(index);
            cell.setCellValue(headername);
            index++;
        }
        int rowCount = 1; 
        while (rs.next()) {
            Row row = sheet.createRow(rowCount);
            for (int i=0; i<numberOfColumns; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(rs.getString(i+1));
            }
            rowCount++;
        }
    }

    
    
    /**
     *
     * @param sheettitle
     * @param rs
     * @throws IOException
     * @throws SQLException
     * 
     * Permet d'enregister le résultat d'une requète SQL dans une feuille EXCEL en conservant les 
     * entêtes de colonnes dans EXCEL identiques à ceux des colonnes de la base SQL 
     */
    public void newSheetFromResultSet_DS(String sheettitle, ResultSet rs, sDSData DSdata ) throws IOException, SQLException {
        
        // Exploitation des ResultSet
//        String sql  = "SELECT * FROM MATABLE"; 
//        Statement statement = connection.createStatement(); 
//        ResultSet resultat = statement.executeQuery(sql); 
//        ResultSetMetaData metadata = resultat.getMetaData(); 
//        int nombreColonnes = metadata.getColumnCount(); 
//        System.out.println("Ce ResultSet contient "+nombreColonnes+" colonnes.");        

        ResultSetMetaData metadata = rs.getMetaData();   // Données complémentaires pour le ResultSet 
        int numberOfColumns = metadata.getColumnCount(); // Nombre de colonnes dans le résultat
        List<String> listHeaderName = new ArrayList<>(numberOfColumns); 
        for(int i = 1; i <= numberOfColumns; i++){ 
            listHeaderName.add(metadata.getColumnName(i)); // index des ResultSet commence à 1
        } 

        if (m_Workbook==null) {
            m_Workbook=getWorkbook();
        }
        Sheet sheet = m_Workbook.createSheet(sheettitle); // On ajoute une feuille au fichier Excel
        int defaultSize = sheet.getColumnWidth(0);
        sheet.setColumnWidth(0, defaultSize); // Numero
        sheet.setColumnWidth(1, defaultSize*3); // Nom
        sheet.setColumnWidth(2, defaultSize*2); // Prénom
        //sheet.setColumnWidth(3, defaultSize*1); // Année
        sheet.setColumnWidth(3, defaultSize*4); // Emargement
        
    // Create a new font and new style for Title and alter it.
        Font fontTitle = m_Workbook.createFont();
        fontTitle.setFontHeightInPoints((short)16);
        //fontTitle.setFontName("Arial");
        //fontTitle.setItalic(true);
        fontTitle.setBold(true);
        //fontTitle.setStrikeout(true);
        CellStyle styleTitle = m_Workbook.createCellStyle();
        styleTitle.setFont(fontTitle);

    // Create a new font and new style for Etudiant and alter it.
        Font fontEtudiants = m_Workbook.createFont();
        fontEtudiants.setFontHeightInPoints((short)11);
        //fontEtudiants.setFontName("Arial");
        //fontEtudiants.setItalic(true);
        //fontEtudiants.setStrikeout(true);
        CellStyle styleEtudiants = m_Workbook.createCellStyle();
        styleEtudiants.setFont(fontEtudiants);

        Map<String, Object> properties = new HashMap<>();
		  
    // border around a cell
        properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
  
    // Give it a color (RED)
//        properties.put(CellUtil.TOP_BORDER_COLOR, IndexedColors.RED.getIndex());
//        properties.put(CellUtil.BOTTOM_BORDER_COLOR, IndexedColors.RED.getIndex());
//        properties.put(CellUtil.LEFT_BORDER_COLOR, IndexedColors.RED.getIndex());
//        properties.put(CellUtil.RIGHT_BORDER_COLOR, IndexedColors.RED.getIndex());
		  
        
        
        int rowCount = 0; 
        int index=1;
        
        Row rowTitle = sheet.createRow(rowCount);
        Cell cell = rowTitle.createCell(index);
        cell.setCellValue("DS : "+DSdata.getNom());
        cell.setCellStyle(styleTitle);
        rowCount++;
        rowTitle = sheet.createRow(rowCount);
        cell = rowTitle.createCell(index);
        cell.setCellValue(DSdata.getDate());        
        cell.setCellStyle(styleTitle);
        rowCount++;
        rowTitle = sheet.createRow(rowCount);
        cell = rowTitle.createCell(index);
        cell.setCellValue("de "+ DSdata.getDebut() + " à "+DSdata.getFin()+"  -  Durée "+ DSdata.getDuree());
        cell.setCellStyle(styleTitle);
        if (DSdata.avecTiersTemps()) {
            rowCount++;
            rowTitle = sheet.createRow(rowCount);
            cell = rowTitle.createCell(index);
            cell.setCellValue("(de "+ DSdata.getDebutTT() + " à "+ DSdata.getFinTT() + " pour les étudiants à temps majoré)");
            cell.setCellStyle(styleTitle);
        }
        rowCount++;
        rowCount++;

        rowTitle = sheet.createRow(rowCount);
        cell = rowTitle.createCell(index);
        cell.setCellValue("Salle "+ DSdata.getSalle());
        cell.setCellStyle(styleTitle);

        rowCount++;
        rowCount++;
        
        // Enregistrement du Header des colommes
        Row rowHeader = sheet.createRow(rowCount);
        index=0;
        for (String headername : listHeaderName) {
            cell = rowHeader.createCell(index);
            cell.setCellValue(headername);
            cell.setCellStyle(styleEtudiants);
            CellUtil.setCellStyleProperties(cell, properties);
            index++;
        }
        cell = rowHeader.createCell(index);
        cell.setCellValue("Emargement");
        cell.setCellStyle(styleEtudiants);
        CellUtil.setCellStyleProperties(cell, properties);
            
        
        rowCount++;

        while (rs.next()) {
            Row row = sheet.createRow(rowCount);
            int defaultHeight=row.getHeight();
            row.setHeight((short) (defaultHeight*2));
            for (int i=0; i<numberOfColumns; i++) {
                cell = row.createCell(i);
                cell.setCellValue(rs.getString(i+1));
                cell.setCellStyle(styleEtudiants);
                CellUtil.setCellStyleProperties(cell, properties);
            }
            // Pour Emargement
            cell = row.createCell(numberOfColumns);
            cell.setCellValue("                            ");
            cell.setCellStyle(styleEtudiants);
            CellUtil.setCellStyleProperties(cell, properties);
            rowCount++;
        }
    }


    
    public void setWorkbook(Workbook workbook) {
        m_Workbook = workbook;
    }
    
    
    
    public void saveAs(String excelFilePath) throws FileNotFoundException, IOException {
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            m_Workbook.write(outputStream);
        }        
    }

    
    
    
    
    
    public static void main(String[] args) throws IOException, InvalidFormatException, sExcelFileManagerException {

        final String SAMPLE_XLS_FILE_PATH = "./test-xls.xls";
        final String SAMPLE_XLSX_FILE_PATH= "./test-xlsx.xlsx";
        
        System.out.println("Working Directory = " + System.getProperty("user.dir"));
        
        System.out.println("\n-----------------------------------------------");
        System.out.println("Test sur le fichier : " + SAMPLE_XLS_FILE_PATH);
        
        //Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        sExcelFileManager sEFM_XLS = new sExcelFileManager(SAMPLE_XLS_FILE_PATH);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + sEFM_XLS.getNumberOfSheets() + " Sheets : ");
        for (String s : sEFM_XLS.getSheetsName()) {
            System.out.println("SheetName " + s);
        }
        
        try {
            List<String> header = sEFM_XLS.getSheetHeader(1); 
            System.out.println("--- Header ----");
            for (String s : header) {
                System.out.println(s);
            }
            System.out.println("---------------");
            
        }
        catch (sExcelFileManagerException e) {
            System.out.println(e.getMessage());
        }

        sExcelFileManager.showSheetData(sEFM_XLS.getSheetFromSheetName("shEeT1")); // Le nom est insensible à la case
        
               
        
        
        System.out.println("\n-----------------------------------------------");
        System.out.println("Test sur le fichier : " + SAMPLE_XLSX_FILE_PATH);
        
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        //Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        sExcelFileManager sEFM_XLSX = new sExcelFileManager(SAMPLE_XLSX_FILE_PATH);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + sEFM_XLS.getNumberOfSheets() + " Sheets : ");
        for (String s : sEFM_XLS.getSheetsName()) {
            System.out.println("SheetName " + s);
        }
        
        try {
            List<String> header = sEFM_XLS.getSheetHeader(7); 
            System.out.println("--- Header ----");
            for (String s : header) {
                System.out.println(s);
            }
            System.out.println("---------------");
            
        }
        catch (sExcelFileManagerException e) {
            System.out.println(e.getMessage());
        }

        sExcelFileManager.showSheetData(sEFM_XLS.getSheetFromSheetName("shEeT1")); // Le nom est insensible à la case
        List<String> listSalles = sEFM_XLS.getColumnFromHeader("Colonne A");
        showListData(listSalles);
        
                
        
    }
    
    
}

