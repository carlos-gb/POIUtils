
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
    
        //Direccion del archivo que se va procesar
        String pathFile = "/home/solrac/pruebas.xls";
        //Cabeceras en el archivo que nos interesa procesar
        String[] cabecera_text = "Cabecera1,Cabecera2,Cabecera3".split(",");

        List cabecera = new ArrayList();
        for (int i = 0; i < cabecera_text.length; i++) {
            //Hacemos a todas las cabeceras opcionales
            cabecera.add(new Header(cabecera_text[i], false));
        }

        ExcelReader reader = new ExcelReader(pathFile, cabecera);
        List<HashMap> hojas = reader.getContent(0);
        List<String[]> filas = (List<String[]>) hojas.get(0).get("data");
        for (int j = 0; j < filas.size(); j++) {
            System.out.println("Fila " + (j + 1) + ": " + arrayToString(filas.get(j)));
        }

    }
    private static String arrayToString(String[] info) {
        String cadena = "";
        for (int i = 0; i < info.length; i++) {
            cadena += info[i] + ",";
        }
        return cadena + "|";
    }
}
