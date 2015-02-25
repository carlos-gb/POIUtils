/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.excel.utils;

import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author adminlx
 */
public class FileProcessing {
    /**
     * Obtiene la extension del archivo
     *
     * @param file ruta del archivo
     * @return String extension
     */
    public static String getExtension(String file){
        String[] tokens=file.split(File.pathSeparator);
        String fileName=tokens[tokens.length-1];
        return fileName.substring(fileName.lastIndexOf(".") + 1);
    }
    /**
     * Genera un objeto WorkBook de un archivo de excel, 
     * sin importar la version del formato
     *
     * @param file ruta del archivo
     * @return WorkBook
     * @throws FileNotFoundException
     * @throws IOException
     */
    public Workbook getObjectData(String file ) throws FileNotFoundException, IOException{
        //InputStream input = new FileInputStream(file);
        Workbook libro=null;
        String extension=this.getExtension(file);
        if(extension.equals(ExcelVersion.v2003)){
            try{
                libro=new HSSFWorkbook(new FileInputStream(file));
            }catch(org.apache.poi.poifs.filesystem.OfficeXmlFileException ex){
                libro=new XSSFWorkbook(new FileInputStream(file));
            }
        }else{
            if(extension.equals(ExcelVersion.v2010)){
                libro=new XSSFWorkbook(new FileInputStream(file));
            }
        }
        return libro;
    }
    
}
