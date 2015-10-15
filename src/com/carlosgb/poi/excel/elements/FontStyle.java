/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.carlosgb.poi.excel.elements;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author carlos
 */
public class FontStyle {
    
    private short color;
    private short size;
    private short boldWeight;
    
    public FontStyle(short color,short size,short boldWeight){
        this.color=color;
        this.size=size;
        this.boldWeight=boldWeight;
    }
    public FontStyle(){
        this.color=IndexedColors.BLACK.index;
        this.size=(short) 12;
        this.boldWeight=Font.BOLDWEIGHT_NORMAL;
    }
    public Font getFont(Workbook wb){
        Font fontStyle=wb.createFont();
        fontStyle.setBoldweight(this.boldWeight);
        fontStyle.setColor(this.color);
        fontStyle.setFontHeightInPoints(this.size);
        return fontStyle;
    }
    /**
     * @return the color
     */
    public short getColor() {
        return color;
    }

    /**
     * @param color the color to set
     */
    public void setColor(short color) {
        this.color = color;
    }

    /**
     * @return the size
     */
    public short getSize() {
        return size;
    }

    /**
     * @param size the size to set
     */
    public void setSize(short size) {
        this.size = size;
    }

    /**
     * @return the boldWeight
     */
    public short getBoldWeight() {
        return boldWeight;
    }

    /**
     * @param boldWeight the boldWeight to set
     */
    public void setBoldWeight(short boldWeight) {
        this.boldWeight = boldWeight;
    }
    
}
