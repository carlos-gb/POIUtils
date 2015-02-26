 /*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.carlosgb.poi.excel.elements;

/**
 *
 * @author carlos
 */
public class Header {
    
    private String nombreColumna;
    private boolean required;
    private int cell;
    private HeaderStyle style;
    private BodyStyle bodyStyle;
    
    public Header(){
        this.nombreColumna=new String();
        this.required=true;
        this.cell=-1;
        this.style=new HeaderStyle();
        this.bodyStyle=new BodyStyle();
    }
    public Header(String nombreColumna){
        this.nombreColumna=nombreColumna;
        this.required=true;
        this.cell=-1;
        this.style=new HeaderStyle();
        this.bodyStyle=new BodyStyle();
    }
    
    public Header(String nombreColumna,boolean required){
        this.nombreColumna=nombreColumna;
        this.required=required;
        this.cell=-1;
        this.style=new HeaderStyle();
        this.bodyStyle=new BodyStyle();
    }
    
    
    /**
     * @return the nombreColumna
     */
    public String getNombreColumna() {
        return nombreColumna;
    }

    /**
     * @param nombreColumna the nombreColumna to set
     */
    public void setNombreColumna(String nombreColumna) {
        this.nombreColumna = nombreColumna;
    }

    /**
     * @return the required
     */
    public boolean isRequired() {
        return required;
    }

    /**
     * @param required the required to set
     */
    public void setRequired(boolean required) {
        this.required = required;
    }

    /**
     * @return the cell
     */
    public int getCell() {
        return cell;
    }

    /**
     * @param cell the cell to set
     */
    public void setCell(int cell) {
        this.cell = cell;
    }

    /**
     * @return the style
     */
    public HeaderStyle getStyle() {
        return style;
    }

    /**
     * @param style the style to set
     */
    public void setStyle(HeaderStyle style) {
        this.style = style;
    }

    /**
     * @return the bodyStyle
     */
    public BodyStyle getBodyStyle() {
        return bodyStyle;
    }

    /**
     * @param bodyStyle the bodyStyle to set
     */
    public void setBodyStyle(BodyStyle bodyStyle) {
        this.bodyStyle = bodyStyle;
    }
    public String toString(){
        return "{nombre_columna:"+this.nombreColumna+", requerido:"+required+"}";
    }

}
