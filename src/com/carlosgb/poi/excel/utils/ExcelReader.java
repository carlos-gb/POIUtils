/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.carlosgb.poi.excel.utils;

import com.carlosgb.poi.excel.elements.Header;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author carlos
 */
public class ExcelReader {

    private String path_file;
    private List<Header> cabecera;
    private Workbook libro;
    private int active_sheet;
    private int header_row;

    public ExcelReader(String path, List<Header> cabecera) {
        this.path_file = path;
        this.cabecera = cabecera;
    }
    public ExcelReader(String path,String[]cabecera){
        this.path_file = path;
        this.cabecera=new ArrayList();
        for(int i=0;i<cabecera.length;i++){
            Header temp=new Header(cabecera[i]);
            this.cabecera.add(temp);
        }
    }

    private void initObject() throws FileNotFoundException, IOException {
        FileProcessing fp = new FileProcessing();
        this.libro = fp.getObjectData(this.path_file);
    }

    public ExcelReader(String path) {
        this.path_file = path;
        this.cabecera = new ArrayList();
    }

    /**
     * Obtiene un List con la informacion de todas las hojas en el archivo de
     * excel y las celdas con las cabeceras especificadas
     *
     * @return HashMap
     * @throws FileNotFoundException
     * @throws IOException
     */
    public List getContent() throws FileNotFoundException, IOException {

        return getContent(null);
    }

    /**
     * Obtiene un List con la informacion de la hoja del archivo de excel
     * indicada en el parametro 'hoja' y las celdas con las cabeceras
     * especificadas
     *
     * @param int hoja
     * @return HashMap
     * @throws FileNotFoundException
     * @throws IOException
     */
    public List getContent(int hoja) throws FileNotFoundException, IOException {
        return getContent(new int[]{hoja});
    }

    /**
     * Obtiene un List con la informacion de las hojas en el parametro
     * 'hojas' y las celdas con las cabeceras especificadas
     *
     * @param int[] hojas
     * @return HashMap
     * @throws FileNotFoundException
     * @throws IOException
     */
    public List getContent(int[] hojas) throws FileNotFoundException, IOException {
        List sheets = new ArrayList();
        if (!this.cabecera.isEmpty()) {
            this.initObject();
            Sheet hoja = libro.getSheetAt(0);
            int numero_de_hojas = 0;
            if (hojas == null) {
                numero_de_hojas = libro.getNumberOfSheets();
            } else {
                numero_de_hojas = hojas.length;
            }
            
            for (int k = 0; k < numero_de_hojas; k++) {
                HashMap sheet = new HashMap();
                List rows_data = new ArrayList();
                this.active_sheet = hojas == null ? k : hojas[k];
                this.header_row=-1;
                hoja = libro.getSheetAt(this.active_sheet);
                if(hoja.getLastRowNum()>0){
                    initHeaders();

                    int to_row = hoja.getLastRowNum() + 1;
                    String[] data_temp = new String[this.cabecera.size()];
                    for (int i = header_row + 1; i < to_row; i++) {
                        data_temp = new String[this.cabecera.size()];
                        for (int z = 0; z < this.cabecera.size(); z++) {
                            if(cabecera.get(z).getCell()>-1){
                                try{   
                                    if(cabecera.get(z).getCell()>-1){
                                        if(!this.cabecera.get(z).getBodyStyle().isIsDate()){
                                            hoja.getRow(i).getCell(cabecera.get(z).getCell()).setCellType(Cell.CELL_TYPE_STRING);
                                            data_temp[z] = hoja.getRow(i).getCell(cabecera.get(z).getCell()).getStringCellValue();
                                        }else if(this.cabecera.get(z).getBodyStyle().isIsDate()){
                                            data_temp[z] = hoja.getRow(i).getCell(cabecera.get(z).getCell()).getDateCellValue().toString();
                                        }
                                    }
                                }catch(NullPointerException ex){
                                    if(!cabecera.get(z).isRequired()){
                                        data_temp[z]=new String();
                                    }
                                }
                            }
                        }
                        //Parche para no meter filas vacias
                        if(!validaFilaVacia(data_temp))
                        rows_data.add(data_temp);
                    }
                    sheet.put("sheet_name", hoja.getSheetName());
                    sheet.put("data", rows_data);
                    sheets.add(sheet);
                }
            }
        } else {
            throw new NullPointerException("No hay cabeceras definidas para la busqueda");
        }
        return sheets;
    }
    /** Metodo para validar filas vacias
     * 
     * @param fila
     * @return <b>true:</b>si la fila esta vacia<br/><b>false:</b>si la fila tiene datos.
     */
    private boolean validaFilaVacia(String[] fila){
        boolean vacia=true;
        for(int i=0;i<fila.length;i++){
            vacia&=fila[i].isEmpty();
        }
        return vacia;
    }
    
    private void initHeaders() {
        Header temp_cabecera = new Header();
        for (int i = 0; i < this.cabecera.size(); i++) {
            temp_cabecera = this.cabecera.get(i);
            if (temp_cabecera.getCell() == -1) {
                if(this.header_row==-1){
                    System.out.println(temp_cabecera.getNombreColumna());
                    int[] temp_xy = this.searchCell(temp_cabecera.getNombreColumna());
                    this.header_row=temp_xy[0];
                    temp_cabecera.setCell(temp_xy[1]);
                }else{
                    temp_cabecera.setCell(searchCell(temp_cabecera.getNombreColumna(), this.header_row));
                }
            }
            
        }
    }

    private int[] searchCell(String text) {
        int[] temp = new int[]{-1, -1};
        text=text.toLowerCase();
        Sheet hoja = libro.getSheetAt(this.active_sheet);
        for (int i = hoja.getFirstRowNum(); i < hoja.getLastRowNum() + 1; i++) {
            Row fila = hoja.getRow(i);
            if(fila!=null){
                for (int j = fila.getFirstCellNum(); j < fila.getLastCellNum(); j++) {
                    try{fila.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
                    if (text.equals(fila.getCell(j).getStringCellValue().toLowerCase())) {
                        temp[0] = i;
                        temp[1] = j;
                        i=hoja.getLastRowNum();
                        break;
                    }}catch(NullPointerException e){}
                }
            }
        }
        return temp;
    }

    private int searchCell(String text, int row_num) {
        int temp = -1;
        Sheet hoja = libro.getSheetAt(this.active_sheet);
        Row fila = hoja.getRow(row_num);
        text=text.toLowerCase();
        for (int j = 0; j < fila.getLastCellNum(); j++) {
            try{
            fila.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
            if (text.equals(fila.getCell(j).getStringCellValue().toLowerCase())) {
                temp=j;
                break;
            }}catch(NullPointerException e){}
        }
        return temp;
    }

    /**
     * @return the path_file
     */
    public String getPathFile() {
        return path_file;
    }

    /**
     * @param path_file the path_file to set
     */
    public void setPathFile(String path_file) {
        this.path_file = path_file;
    }

    /**
     * @return the cabecera
     */
    public List<Header> getCabecera() {
        return cabecera;
    }

    /**
     * @param cabecera the cabecera to set
     */
    public void setCabecera(List<Header> cabecera) {
        this.cabecera = cabecera;
    }
}
