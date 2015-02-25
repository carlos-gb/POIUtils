
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProtectedExcelFile {

    public static void main(final String... args) throws Exception {

        String fname = "/home/adminlx/sample2.xlsx";

        FileOutputStream fileOut = null;

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet();

            XSSFRow row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            sheet.protectSheet("webmaster");
            cell.setCellValue("THIS WORKS!");

            fileOut = new FileOutputStream(fname);
            workbook.write(fileOut);

        } catch (Exception ex) {

            System.out.println(ex.getMessage());

        } finally {

            try {

                fileOut.close();

            } catch (IOException ex) {

                System.out.println(ex.getMessage());

            }
        }

    }
}
