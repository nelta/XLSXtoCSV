import java.io.*;
import java.nio.charset.Charset;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class XlsxtoCSV {

    private static void xlsx(File inputFile, File outputFile) {

        final Charset ASCII = Charset.forName("ASCII");

        // For storing data into CSV files
        StringBuilder data = new StringBuilder();
        try {
            Writer fos = new OutputStreamWriter(new FileOutputStream(outputFile), ASCII);

            // Get the workbook object for XLSX file
            XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(inputFile));

            // Get first sheet from the workbook
            XSSFSheet sheet = wBook.getSheetAt(0);
            Row row;
            Cell cell;

            // Iterate through each row from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            //append header
            data.append("PersNr,Vorname,Nachname,Datum,Kommen,Gehen,Pause\r\n");

            //select first row
            row = rowIterator.next();


            while (rowIterator.hasNext()) {
                while (row.getCell(0)== null || !(row.getCell(0).getStringCellValue().equals("Person"))){
                    row = rowIterator.next();
                }
                //get PersNr from E5
                String persNr = (int) row.getCell(4).getNumericCellValue() + ",";
                row = rowIterator.next();
                //get full name from F6
                String name = row.getCell(5).getStringCellValue();
                //split full name
                String prename = name.substring(name.indexOf(",") + 2) + ",";
                String surname = name.substring(0, name.indexOf(",") + 1);

                //skip irrelevant rows
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                String curDate = "";

                while (rowIterator.hasNext()) {
                    row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    cell = cellIterator.next();
                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        if (cell.getStringCellValue().equals("Monatssumme")) break;
                        if (cell.getStringCellValue().equals("Wochensumme")) continue;
                    }

                    //append final values for this worker
                    data.append(persNr).append(prename).append(surname);

                    //append date
                    if (cell.getColumnIndex() == 0) {
                        double date = cell.getNumericCellValue();
                        curDate = (date < 10 ? "0" + date : "" + date) + ".";
                    }
                    data.append(curDate);

                    //append "Kommen"
                    cell = row.getCell(8);
                    if (cell == null) {
                        data.append(",,,\r\n");
                        continue;
                    }
                    data.append(cell.getStringCellValue().trim()).append(",");
                    //append "Gehen"
                    cell = row.getCell(10);
                    data.append(cell.getStringCellValue().trim()).append(",");
                    //append "Pause"
                    cell = row.getCell(12);
                    data.append(cell == null ? "0:00" : cell.getStringCellValue().trim());

                    data.append("\r\n");
                    System.out.println(row.getRowNum());
                }

                //skip irrelevant rows
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();

                System.out.println(row.getRowNum());
            }

            fos.write(data.toString());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    public static void main(String[] args) {
        File inputFile = new File("C:\\Users\\Nelta\\Desktop\\Stryker_Zeiterfassung.xlsx");
        File outputFile = new File("C:\\Users\\Nelta\\Desktop\\output.csv");
        xlsx(inputFile, outputFile);
    }
}