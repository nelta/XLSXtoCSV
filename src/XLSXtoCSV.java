import java.io.*;
import java.nio.charset.Charset;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class XLSXtoCSV {

    private static void convertXLSX(File inputFile, File outputFile) {

        //charset required by Landwehr
        final Charset ASCII = Charset.forName("ASCII");

        //Stringbuilder for storing data into CSV files
        StringBuilder data = new StringBuilder();
        try {
            Writer fos = new OutputStreamWriter(new FileOutputStream(outputFile), ASCII);

            // Get the workbook object for XLSX file
            XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(inputFile));

            // Get first sheet from the workbook
            XSSFSheet sheet = wBook.getSheetAt(0);
            Row row;
            Cell cell;

            //Iterator to iterate through each row from sheet
            Iterator<Row> rowIterator = sheet.iterator();
            //append predefined header
            data.append("PersNr,Vorname,Nachname,Datum,Kommen,Gehen,Pause\r\n");

            //select first row
            row = rowIterator.next();

            outerLoop:
            while (rowIterator.hasNext()) {
                //skip rows until the first cell contains the word "Person"
                while (row.getCell(0)== null || !(row.getCell(0).getStringCellValue().equals("Person"))){
                    //stop when no rows left
                    if (rowIterator.hasNext()) row = rowIterator.next();
                    else break outerLoop;
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
                // TODO: 21.07.2017 design it dynamically, maybe Hydra-Export is not consistent
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();
                row = rowIterator.next();

                //define the current date, used when multiple records have been made for one day
                String curDate = "";

                while (rowIterator.hasNext()) {
                    //select first row with records
                    row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    //select first cell of current row
                    cell = cellIterator.next();
                    //when the first cell in a row contains the word "Monatssumme" that employees record have been
                    //completed and the inner while loop is escaped
                    //when the first cell in a row contains the word "Wochensumme" that row can be skipped
                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        if (cell.getStringCellValue().equals("Monatssumme")) break;
                        if (cell.getStringCellValue().equals("Wochensumme")) continue;
                    }

                    //append final values for this worker (in every record)
                    data.append(persNr).append(prename).append(surname);

                    //append date, when the first selectable cells index in that row is not 0, use preciously stored
                    //current date, else read the cell
                    if (cell.getColumnIndex() == 0) {
                        double date = cell.getNumericCellValue();
                        //fit date format to requirements (dd.mm.)
                        curDate = (date < 10 ? "0" + date : "" + date) + ".";
                    }
                    data.append(curDate);

                    //append "Kommen"
                    cell = row.getCell(8);
                    //when no record have been made for a certain date, append corresponding commas for the empty
                    //entries and a linebreak, skip all other commands in this loop
                    if (cell == null) {
                        data.append(",,,\r\n");
                        continue;
                    }
                    //some cell contain a leading whitespace, so trim is needed
                    data.append(cell.getStringCellValue().trim()).append(",");

                    //append "Gehen"
                    cell = row.getCell(10);
                    data.append(cell.getStringCellValue().trim()).append(",");

                    //append "Pause", multiple entries on one day do have one entry for "Pause", add "0:00" instead
                    cell = row.getCell(12);
                    data.append(cell == null ? "0:00" : cell.getStringCellValue().trim());

                    //append final linebreak for that record
                    data.append("\r\n");
                }
            }

            //write everything to output file
            fos.write(data.toString());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    public static void main(String[] args) {
        File inputFile = new File("C:\\Users\\Nelta\\Desktop\\Stryker_Zeiterfassung.convertXLSX");
        File outputFile = new File("C:\\Users\\Nelta\\Desktop\\output.txt");
        convertXLSX(inputFile, outputFile);
    }
}