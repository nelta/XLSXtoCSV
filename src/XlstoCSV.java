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

                        // TODO: 20.07.2017 while starts here for workers

                        //select first row
                        rowIterator.next();
                        //skip irrelevant row
                        row = rowIterator.next();
                        //get PersNr from E5
                        String persNr = (int)row.getCell(4).getNumericCellValue() + ",";
                        row = rowIterator.next();
                        //get full name from F6
                        String name = row.getCell(5).getStringCellValue();
                        //split full name
                        String prename = name.substring(name.indexOf(",")+2) + ",";
                        String surname = name.substring(0,name.indexOf(",")+1);

                        //skip irrelevant rows
                        row = rowIterator.next();
                        row = rowIterator.next();
                        row = rowIterator.next();
                        row = rowIterator.next();
                        String curDate = "";

                        while ( rowIterator.hasNext()) {
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
                                if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC) {
                                        double date = cell.getNumericCellValue();
                                        curDate = (date < 10 ? "0" + date : "" + date) + ".";
                                        //skip one column
                                        cell = cellIterator.next();
                                        if (cellIterator.hasNext()) cell = cellIterator.next();
                                }
                                data.append(curDate);


                                if (cellIterator.hasNext()) {
                                        //append "Kommen"
                                        data.append(",").append(cell.getStringCellValue()).append(",");
                                        cell = cellIterator.next();
                                        //append "Gehen"
                                        data.append(cell.getStringCellValue()).append(",");
                                        cell = cellIterator.next();
                                        //append "Pause"
                                        data.append(cell.getStringCellValue());
                                } else {
                                        data.append(",,,");
                                }
                                data.append("\r\n");
                                System.out.println("*");
                        }



                        /*while (rowIterator.hasNext()) {
                                row = rowIterator.next();

                                // For each row, iterate through each columns
                                Iterator<Cell> cellIterator = row.cellIterator();
                                while (cellIterator.hasNext()) {

                                        cell = cellIterator.next();

                                        switch (cell.getCellType()) {
                                                case Cell.CELL_TYPE_BOOLEAN:
                                                        data.append(cell.getBooleanCellValue() + (cellIterator.hasNext() ? "," : ""));

                                                        break;
                                                case Cell.CELL_TYPE_NUMERIC:
                                                        data.append(cell.getNumericCellValue() + (cellIterator.hasNext() ? "," : ""));

                                                        break;
                                                case Cell.CELL_TYPE_STRING:
                                                        data.append(cell.getStringCellValue() + (cellIterator.hasNext() ? "," : ""));
                                                        break;

                                                case Cell.CELL_TYPE_BLANK:
                                                        data.append("" + (cellIterator.hasNext() ? "," : ""));
                                                        break;
                                                default:
                                                        data.append(cell + (cellIterator.hasNext() ? "," : ""));

                                        }
                                }
                                data.append("\r\n");
                        }*/

                        fos.write(data.toString());
                        fos.close();

                } catch (Exception ioe) {
                        ioe.printStackTrace();
                }
        }

        public static void main(String[] args) {
                File inputFile = new File("C:\\Users\\Nelta\\Desktop\\Stryker_Zeiterfassung.xlsx");
                File outputFile = new File("C:\\Users\\Nelta\\Desktop\\output.txt");
                xlsx(inputFile, outputFile);
        }
}