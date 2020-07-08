import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) {
		try {
			File file = new File("D:\\sheet.xlsx"); // creating a new file instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
			String tab = "\t";
			String nl = "\n";
			String toSend = "";
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			wb.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);
			System.out.println("Sheets: " + wb.getNumberOfSheets());
			DataFormatter formatter = new DataFormatter();
			XSSFSheet sheet = wb.getSheetAt(0);

			for (int rn = sheet.getFirstRowNum(); rn <= sheet.getLastRowNum(); rn++) {

				Row row = sheet.getRow(rn);
				if (row == null) {

					// There is no data in this row, handle as needed
				} else {

					for (int cn = 0; cn < row.getLastCellNum(); cn++) {
						Cell cell = row.getCell(cn);

						if (cell == null) {

							String text = "";
							toSend = toSend + text + tab;

						} else {
							String text = formatter.formatCellValue(cell);
							toSend = toSend + text + tab;

						}
					}

					toSend = toSend + nl;

				}

			}

			System.out.print(toSend);

			wb.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
