
import com.excel.elements.Header;
import com.excel.utils.ExcelReader;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/*
 * To change this template, choose Tools | Templates and open the template in
 * the editor.
 */
/**
 *
 * @author Carlos
 */
public class LeeExcel {

    public static void main(String[] args) throws FileNotFoundException, IOException {

        String pathFile = "/home/adminlx/Documents/Hixsa Files/CEFIDI/cuentas SAT/importacion cuentas - prueba.xls";
        //String[] cabecera_text = new String[]{"Nro. ID", "Nombre", "Fecha/Hora", "Terminal N º", "nombre del dispositivo",
        //    "Estado", "Descripción Estado", "Sector", "Cargo"};

        String[] cabecera_text = "Código del servicio,Nombre del servicio,Impuesto 1,Precio 1".split(",");

        List cabecera = new ArrayList();
        for (int i = 0; i < cabecera_text.length; i++) {
            cabecera.add(new Header(cabecera_text[i], false));
        }

        ExcelReader reader = new ExcelReader(pathFile, cabecera);
        List<HashMap> hojas = reader.getContent(0);
        List<String[]> filas = (List<String[]>) hojas.get(0).get("data");
        for (int j = 0; j < filas.size(); j++) {
            //System.out.println("Fila " + (j + 1) + ": " + arrayToString(filas.get(j)));
            String cadenaReducida = limitarCadena(filas.get(j)[1], 50);
            System.out.println("Fila " + (j + 1) + " Cadena reducida:" + cadenaReducida.length() + " - " + cadenaReducida);
            if(filas.get(j)[1].length()>50)
            System.out.println("Restante:" + filas.get(j)[1].replace(cadenaReducida, ""));

        }

    }

    private static String limitarCadena(String cadena, int maxLenght) {
        String cadenaLimit = new String();
        String[] palabras = cadena.split(" ");
        if (cadena.length() > maxLenght) {
            for (int i = 0; i < palabras.length; i++) {
                int tempLength = cadenaLimit.length()
                        + palabras[i].length();
                if (tempLength < maxLenght) {
                    cadenaLimit += " " + palabras[i];
                }else{
                    break;
                }
            }
            if (!cadenaLimit.isEmpty()) {
                cadenaLimit = cadenaLimit.substring(1);
            }
        } else {
            cadenaLimit = cadena;
        }
        return cadenaLimit;
    }

    private static String arrayToString(String[] info) {
        String cadena = "";
        for (int i = 0; i < info.length; i++) {
            cadena += info[i] + ",";
        }
        return cadena + "|";
    }
}
