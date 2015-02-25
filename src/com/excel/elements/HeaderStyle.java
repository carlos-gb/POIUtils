/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.excel.elements;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author carlos
 */
public class HeaderStyle {
    
    private short textAlign;
    private short verticalAlign;
    private short borderRight;
    private short borderLeft;
    private short borderTop;
    private short borderBottom;
    private short backgroundColor;
    private boolean wrapText;
    private boolean customBackgroundColor;
    private FontStyle fontStyle;
    
    private byte redBackgroundColor;
    private byte greenBackgroundColor;
    private byte blueBackgroundColor;
    
    public HeaderStyle(){
        this.fontStyle=new FontStyle(IndexedColors.WHITE.index,(short)14,Font.BOLDWEIGHT_BOLD);
        this.wrapText=false;
        this.backgroundColor=HSSFColor.GREY_40_PERCENT.index;
        this.textAlign=CellStyle.ALIGN_CENTER;
        this.verticalAlign=CellStyle.VERTICAL_CENTER;
        this.borderRight=CellStyle.BORDER_THIN;
        this.borderLeft=CellStyle.BORDER_THIN;
        this.borderTop=CellStyle.BORDER_THIN;
        this.borderBottom=CellStyle.BORDER_THIN;
        this.customBackgroundColor=true;
        this.redBackgroundColor=25;
        this.greenBackgroundColor=91;
        this.blueBackgroundColor=119;
    }
    
    public CellStyle getCellStyle(Workbook wb){
        CellStyle cellStyle=wb.createCellStyle();
        cellStyle.setFont(this.fontStyle.getFont(wb));
        cellStyle.setAlignment(this.textAlign);
        
        cellStyle.setVerticalAlignment(verticalAlign);
        if(this.borderBottom!=-1){
        cellStyle.setBorderBottom(this.borderBottom);}
        if(this.borderLeft!=-1){
        cellStyle.setBorderLeft(this.borderLeft);}
        if(this.borderRight!=-1){
        cellStyle.setBorderRight(this.borderRight);}
        if(this.borderTop!=-1){
        cellStyle.setBorderTop(this.borderTop);}
        cellStyle.setWrapText(this.wrapText);
        if(this.customBackgroundColor){
            this.setForegroundColor(wb);
        }
        cellStyle.setFillForegroundColor(this.getBackgroundColor());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        
        return cellStyle;
    }
    private void setForegroundColor(Workbook wb){
        try{
            HSSFPalette palette=((HSSFWorkbook)wb).getCustomPalette();
            palette.setColorAtIndex(this.getBackgroundColor(), this.getRedBackgroundColor(), this.getGreenBackgroundColor(), this.getBlueBackgroundColor());
        }catch(Exception ex){
            /* byte[]rgb=new byte[]{this.redBackgroundColor,this.greenBackgroundColor,this.blueBackgroundColor};
            XSSFColor color=new XSSFColor(rgb);
            this.setBackgroundColor(color.getIndexed());*/
        }
    }
    /*
     * @return the textAlign
     */
    public short getTextAlign() {
        return textAlign;
    }

    /**
     * @param textAlign the textAlign to set
     */
    public void setTextAlign(short textAlign) {
        this.textAlign = textAlign;
    }

    /**
     * @return the verticalAlign
     */
    public short getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * @param verticalAlign the verticalAlign to set
     */
    public void setVerticalAlign(short verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    /**
     * @return the borderRight
     */
    public short getBorderRight() {
        return borderRight;
    }

    /**
     * @param borderRight the borderRight to set
     */
    public void setBorderRight(short borderRight) {
        this.borderRight = borderRight;
    }

    /**
     * @return the borderLeft
     */
    public short getBorderLeft() {
        return borderLeft;
    }

    /**
     * @param borderLeft the borderLeft to set
     */
    public void setBorderLeft(short borderLeft) {
        this.borderLeft = borderLeft;
    }

    /**
     * @return the borderTop
     */
    public short getBorderTop() {
        return borderTop;
    }

    /**
     * @param borderTop the borderTop to set
     */
    public void setBorderTop(short borderTop) {
        this.borderTop = borderTop;
    }

    /**
     * @return the borderBottom
     */
    public short getBorderBottom() {
        return borderBottom;
    }

    /**
     * @param borderBottom the borderBottom to set
     */
    public void setBorderBottom(short borderBottom) {
        this.borderBottom = borderBottom;
    }

    /**
     * @return the wrapText
     */
    public boolean isWrapText() {
        return wrapText;
    }

    /**
     * @param wrapText the wrapText to set
     */
    public void setWrapText(boolean wrapText) {
        this.wrapText = wrapText;
    }

    /**
     * @return the fontStyle
     */
    public FontStyle getFontStyle() {
        return fontStyle;
    }

    /**
     * @param fontStyle the fontStyle to set
     */
    public void setFontStyle(FontStyle fontStyle) {
        this.fontStyle = fontStyle;
    }

    /**
     * @return the redBackgroundColor
     */
    public byte getRedBackgroundColor() {
        return redBackgroundColor;
    }

    /**
     * @param redBackgroundColor the redBackgroundColor to set
     */
    public void setRedBackgroundColor(byte redBackgroundColor) {
        this.redBackgroundColor = redBackgroundColor;
    }

    /**
     * @return the greenBackgroundColor
     */
    public byte getGreenBackgroundColor() {
        return greenBackgroundColor;
    }

    /**
     * @param greenBackgroundColor the greenBackgroundColor to set
     */
    public void setGreenBackgroundColor(byte greenBackgroundColor) {
        this.greenBackgroundColor = greenBackgroundColor;
    }

    /**
     * @return the blueBackgroundColor
     */
    public byte getBlueBackgroundColor() {
        return blueBackgroundColor;
    }

    /**
     * @param blueBackgroundColor the blueBackgroundColor to set
     */
    public void setBlueBackgroundColor(byte blueBackgroundColor) {
        this.blueBackgroundColor = blueBackgroundColor;
    }

    /**
     * @return the backgroundColor
     */
    public short getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * @param backgroundColor the backgroundColor to set
     */
    public void setBackgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    /**
     * @return the customForegroundColor
     */
    public boolean isCustomBackgroundColor() {
        return customBackgroundColor;
    }

    /**
     * @param customForegroundColor the customForegroundColor to set
     */
    public void setCustomBackgroundColor(boolean customForegroundColor) {
        this.customBackgroundColor = customForegroundColor;
    }

}
