/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.carlosgb.poi.excel.utils;

import com.carlosgb.poi.excel.elements.DataPoiRow;
import com.carlosgb.poi.excel.elements.Header;
import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hpsf.WritingNotSupportedException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author carlos
 */
public class ExcelExport {

    private String path;
    private String fileName;
    private String version;
    private String autor = new String();

    public ExcelExport() {
        this.path = File.separator;
        this.fileName = "export";
        this.version = ExcelVersion.v2003;
    }

    public ExcelExport(String path, String fileName, String version) {
        this.path = validatePath(path);
        this.fileName = fileName;
        this.version = this.validateVersion(version);

    }

    public String createFile(HashMap sheet) throws FileNotFoundException, IOException {
        String full_path = this.path + this.fileName + "." + this.version;
        File temp = new File(this.path);
        if (!temp.exists()) {
            temp.mkdirs();
        }
        //FileOutputStream out = new FileOutputStream(full_path);
        Workbook wb = getNewObject();
        sheet.put("sheet_number", 0);
        wb = this.createSheet(wb, sheet);
        //wb.write(out);
        this.writeWorkbook(wb, full_path);
        return full_path;
    }

    private boolean writeWorkbook(Workbook wb, String full_path) throws FileNotFoundException, IOException {
        FileOutputStream out = new FileOutputStream(full_path);
        wb.write(out);
        out.close();
        if (!this.autor.isEmpty()&&this.version.equals(ExcelVersion.v2003)) {
            File poiFilesystem = new File(full_path);
            InputStream is = new FileInputStream(poiFilesystem);
            POIFSFileSystem poifs = new POIFSFileSystem(is);
            is.close();

            DirectoryEntry dir = poifs.getRoot();
            SummaryInformation si = null;
            try {
                DocumentEntry siEntry = (DocumentEntry) dir.getEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                DocumentInputStream dis = new DocumentInputStream(siEntry);
                PropertySet ps = new PropertySet(dis);
                dis.close();
                si = new SummaryInformation(ps);
            } catch (Exception ex){
                ex.printStackTrace();
                si = PropertySetFactory.newSummaryInformation();
            }
            try {
                si.setAuthor(autor);
                si.write(dir, SummaryInformation.DEFAULT_STREAM_NAME);
            } catch (WritingNotSupportedException ex) {
                ex.printStackTrace();
            }
            OutputStream outStream = new FileOutputStream(poiFilesystem);
            poifs.writeFilesystem(outStream);
            outStream.close();
        }

        return true;
    }

    public String createFile(List sheet) throws FileNotFoundException, IOException {
        String full_path = this.path + this.fileName + "." + this.version;
        File temp = new File(this.path);
        if (!temp.exists()) {
            temp.mkdirs();
        }
        //FileOutputStream out = new FileOutputStream(full_path);
        Workbook wb = getNewObject();
        for (int i = 0; i < sheet.size(); i++) {
            HashMap sheet_map = (HashMap) sheet.get(i);
            sheet_map.put("sheet_number", i);
            wb = this.createSheet(wb, sheet_map);
        }
        //wb.write(out);
        this.writeWorkbook(wb, full_path);
        return full_path;
    }

    private Workbook createSheet(Workbook wb, HashMap sheet_map) {
        Sheet hoja;
        Row row;
        float rowHeightHeader = 30;
        float rowHeightBody = 20;

        if (sheet_map.get("header_height") != null) {
            rowHeightHeader = (Float)sheet_map.get("header_height");
        }
        if (sheet_map.get("body_height") != null) {
            rowHeightBody = (Float) sheet_map.get("body_height");
        }
        
        hoja = wb.createSheet();
        wb.setSheetName((Integer) sheet_map.get("sheet_number"), sheet_map.get("sheet_name") != null ? (String) sheet_map.get("sheet_name") : (Integer) sheet_map.get("sheet_number") + "");
        List<Header> headers = new ArrayList();
        String[] temp_cabecera =null;
        if (sheet_map.get("header").getClass().getName().equals("java.util.ArrayList")) {
            headers = (List) sheet_map.get("header");
            temp_cabecera=new String[headers.size()];
            for (int z=0;z<headers.size();z++) {
                temp_cabecera[z]=headers.get(z).getNombreColumna();
            }
        } else {
            temp_cabecera = (String[]) sheet_map.get("header");
            for (int p = 0; p < temp_cabecera.length; p++) {
                Header tempHeaderObj = new Header(temp_cabecera[p]);
                headers.add(tempHeaderObj);
            }
        }
        List rows = (List) sheet_map.get("data");
        
        //****Para android
        //int size_columns[]=new int[headers.size()];
        
        int index_row = 0;
        row = hoja.createRow(index_row++);
        CellStyle[] bodyStyles = new CellStyle[headers.size()];
        for (int z = 0; z < headers.size(); z++) {
            //****Para android
            //size_columns[z]=headers.get(z).getNombreColumna().length();
            insertData(row, z, headers.get(z).getNombreColumna(), headers.get(z).getStyle().getCellStyle(wb));
            bodyStyles[z] = headers.get(z).getBodyStyle().getCellStyle(wb);
            row.setHeightInPoints(rowHeightHeader);
        }
        for (int j = 0; j < rows.size(); j++) {
            row = hoja.createRow(index_row++);
            row.setHeightInPoints(rowHeightBody);
            String[] data_temp = null;
            if(rows.get(j).getClass().isArray()){
                data_temp=(String[])rows.get(j);
            }else if(rows.get(j) instanceof DataPoiRow){
                data_temp=new String[headers.size()];
                HashMap<String,Object> dataRowtemp=((DataPoiRow)rows.get(j)).toHashMapData();
                int k=0;
                for (String key : temp_cabecera) {
                    data_temp[k++]=dataRowtemp.get(key).toString();
                }
            }
            for (int k = 0; k < data_temp.length; k++) {
                if (!headers.get(k).isRequired() ? true : !headers.get(k).isRequired() || !data_temp[k].isEmpty()) {
                    //****Para android
                    //int tempSize=data_temp[k].length();
                    //if(tempSize>size_columns[k]){size_columns[k]=tempSize;}
                    if (headers.get(k).getBodyStyle().isDouble() && !data_temp[k].isEmpty()) {
                        insertData(row, k, (new BigDecimal(data_temp[k])).doubleValue(), bodyStyles[k]);
                    } else {
                        insertData(row, k, data_temp[k], bodyStyles[k]);
                    }
                } else {
                    throw new NullPointerException("Falta el dato '" + headers.get(k).getNombreColumna() + "' requerido en el registro " + (j + 1) + ".");
                }
            }
        }
        for (int w = 0; w < headers.size(); w++) {
            hoja.autoSizeColumn(w);
            //****Para android
            /*int widthColumn=(size_columns[w]+5)*256;
            if(size_columns[w]>255){
                widthColumn=255*256;
            }
            hoja.setColumnWidth(w,widthColumn);*/
        }
        //protegemos la hoja en caso de que manden una contraseña como parametro en el HashMap
        if (sheet_map.containsKey("password")) {
            hoja.protectSheet(sheet_map.get("password").toString());
        }

        return wb;
    }

    private Row insertData(Row fila, int cell, Object text, CellStyle cellStyle) {
        Cell celda = fila.createCell(cell);
        if (text.getClass().getName().equals((new String()).getClass().getName())) {
            celda.setCellValue((String) text);
        } else {
            celda.setCellValue((Double) text);
        }
        celda.setCellStyle(cellStyle);
        return fila;
    }

    private Workbook getNewObject() {
        Workbook libro = null;
        if (this.version.equals(ExcelVersion.v2003)) {
            libro = new HSSFWorkbook();
        } else {
            libro = new XSSFWorkbook();
        }
        return libro;
    }

    private String validatePath(String path) {
        if (!path.endsWith(File.separator)) {
            path += File.separator;
        }
        return path;
    }

    private String validateVersion(String version) {
        if (!version.equals(ExcelVersion.v2003) && !version.equals(ExcelVersion.v2010)) {
            this.setVersion(ExcelVersion.v2003);
        }
        return version;
    }

    /**
     * @return the path
     */
    public String getPath() {
        return path;
    }

    /**
     * @param path the path to set
     */
    public void setPath(String path) {
        this.path = validatePath(path);
    }

    /**
     * @return the fileName
     */
    public String getFileName() {
        return fileName;
    }

    /**
     * @param fileName the fileName to set
     */
    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    /**
     * @return the version
     */
    public String getVersion() {
        return version;
    }

    /**
     * @param version the version to set
     */
    public void setVersion(String version) {
        this.version = validateVersion(version);
    }

    /**
     * @param autor the autor to set
     */
    public void setAutor(String autor) {
        this.autor = autor;
    }
}
