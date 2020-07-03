import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) {
		try {
			File file = new File("D:\\sheet.xlsx"); // creating a new file instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
//creating Workbook instance that refers to .xlsx file
			String tab = "\t";
			String nl = "\n";
			String toSend = "";
			String suffix = "";
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			System.out.println("Sheets: " + wb.getNumberOfSheets());
			// wb.getNumberOfSheets();
				XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
				Iterator<Row> itr = sheet.iterator(); // iterating over excel file
				int noOfColumns = sheet.getRow(0).getLastCellNum();
				System.out.println("Columns: " + noOfColumns);
				DataFormatter formatter = new DataFormatter();
				while (itr.hasNext()) {
					Row row = itr.next();
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						int cellNum = cell.getColumnIndex();
						if (cellNum == noOfColumns - 1) {
							suffix = nl;
						}

						else {

							suffix = tab;
						}

						String text = formatter.formatCellValue(cell);
						toSend = toSend + text + suffix;

					}
					// System.out.println("");

				}
				System.out.print(toSend);
			

			wb.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}