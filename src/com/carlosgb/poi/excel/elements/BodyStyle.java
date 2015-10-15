/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.carlosgb.poi.excel.elements;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author carlos
 */
public class BodyStyle extends HeaderStyle {
        
    private String format;
    private boolean isDouble;
    private boolean isDate;
    
    public BodyStyle(){
        super();
        initBodyStyle();
    }
    public BodyStyle(String format){
        super();
        initBodyStyle();
        this.format=format;
    }
    private void initBodyStyle(){
        this.setIsDouble(false);
        this.setFontStyle(new FontStyle());
        this.setBackgroundColor(HSSFColor.WHITE.index);
        this.setTextAlign(CellStyle.ALIGN_LEFT);
        this.setCustomBackgroundColor(false);
        this.setRedBackgroundColor((byte)0);
        this.setGreenBackgroundColor((byte)0);
        this.setBlueBackgroundColor((byte)0);
        this.format=new String();
        this.setIsDate(false);
    }
    /**
     * @return the format
     */
    public String getFormat() {
        return format;
    }

    /**
     * @param format the format to set
     */
    public void setFormat(String format) {
        this.format = format;
    }
     public CellStyle getCellStyle(Workbook wb){
        CellStyle cellStyle=super.getCellStyle(wb);
        if(!this.format.isEmpty()){
            DataFormat dataFormat = wb.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat(this.format));
        }
        return cellStyle;
    }

    /**
     * @return the isDouble
     */
    public boolean isDouble() {
        return isDouble;
    }

    /**
     * @param isDouble the isDouble to set
     */
    public void setIsDouble(boolean isDouble) {
        this.isDouble = isDouble;
    }

    public boolean isIsDate() {
        return isDate;
    }

    public void setIsDate(boolean isDate) {
        this.isDate = isDate;
    }
}
