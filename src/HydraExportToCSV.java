import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Iterator;

class HydraExportToCSV {
    //some statistics
    private static int employeesFound = 0;
    private static int recordsFound = 0;

    //charset required by Landwehr
    private final static Charset ASCII = Charset.forName("ASCII");
    //charset for logfile
    private final static Charset UTF8 = Charset.forName("UTF8");

    //List of extracted records
    private static ArrayList<MetaRecord> metaRecords = new ArrayList<>();

    private static void convert(File inputFile, File outputFile, String fileExtension) {

        //extract records
        extractRecords(inputFile, fileExtension);
        for (MetaRecord m :
                metaRecords) {
            System.out.println(m);
        }
        //convert records and write to new file

        //create logfile

        //Stringbuilder for storing data into CSV files
        StringBuilder data = new StringBuilder();
        try {
            Writer fos = new OutputStreamWriter(new FileOutputStream(outputFile), ASCII);

            // Get the workbook object for XLSX or XLS file
            Workbook wBook;
            if (fileExtension.equals("xlsx")) wBook = new XSSFWorkbook(new FileInputStream(inputFile));
            else wBook = new HSSFWorkbook(new FileInputStream(inputFile));

            // Get first sheet from the workbook
            Sheet sheet = wBook.getSheetAt(0);

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
                while (row.getCell(0) == null || !(row.getCell(0).getStringCellValue().equals("Person"))) {
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

                //skip irrelevant rows, select row with "Datum" in first cell
                while (row.getCell(0) == null || !row.getCell(0).getStringCellValue().equals("Datum")) {
                    row = rowIterator.next();
                }

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

                    //append final values for this employee (in every record)
                    data.append(persNr).append(prename).append(surname);

                    //append date, when the first selectable cell's index in that row is not 0, use preciously stored
                    //current date, else read the cell
                    if (cell.getColumnIndex() == 0) {
                        double date = cell.getNumericCellValue();
                        //fit date format to requirements (dd.mm.)
                        curDate = (date < 10 ? "0" + date : "" + date) + ".";
                    }
                    data.append(curDate).append(",");

                    //append "Kommen"
                    cell = row.getCell(8);
                    //when no record have been made for a certain date, append corresponding commas for the empty
                    //entries and a linebreak, skip all other commands in this loop
                    if (cell == null) {
                        data.append(",,\r\n");
                        continue;
                    }
                    //some cells contain a leading whitespace and need to be trimmed
                    data.append(cell.getStringCellValue().trim()).append(",");

                    //append "Gehen"
                    cell = row.getCell(10);
                    data.append(cell.getStringCellValue().trim()).append(",");

                    //append "Pause", multiple entries on one day do have one entry for "Pause", add "0:00" instead
                    cell = row.getCell(12);
                    data.append(cell == null ? "0:00" : cell.getStringCellValue().trim());

                    //append final linebreak for that record
                    data.append("\r\n");

                    //count total records
                    recordsFound++;
                }
                //count total employees
                employeesFound++;
            }

            //write everything to output file
            fos.write(data.toString());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ein Fehler ist aufgetreten. \n" +
                    "Bitten wenden Sie sich an einen Administrator.\n\n Nachricht:\n " + ioe.getMessage());
        }
    }

    private static void extractRecords(File inputFile, String fileExtension) {
        try {
            // Get the workbook object for XLSX or XLS file
            Workbook wBook;
            if (fileExtension.equals("xlsx")) wBook = new XSSFWorkbook(new FileInputStream(inputFile));
            else wBook = new HSSFWorkbook(new FileInputStream(inputFile));

            // Get first sheet from the workbook
            Sheet sheet = wBook.getSheetAt(0);

            Row row;
            Cell cell;

            //Iterator to iterate through each row from sheet
            Iterator<Row> rowIterator = sheet.iterator();

            //select first row
            row = rowIterator.next();

            outerLoop:
            while (rowIterator.hasNext()) {
                //skip rows until the first cell contains the word "Person"
                while (row.getCell(0) == null || !(row.getCell(0).getStringCellValue().equals("Person"))) {
                    //stop when no rows left
                    if (rowIterator.hasNext()) row = rowIterator.next();
                    else break outerLoop;
                }
                //get PersNr(= id = Transpondernummer) from E (column 5)
                int id = (int) row.getCell(4).getNumericCellValue();

                int month;
                int year;
                //get month and year depending on backslash
                if ( row.getCell(39) != null && row.getCell(39).getCellType() == Cell.CELL_TYPE_STRING) {
                    month = (int) row.getCell(38).getNumericCellValue();
                    year = (int) row.getCell(41).getNumericCellValue();
                } else if (row.getCell(38) != null && row.getCell(38).getCellType() == Cell.CELL_TYPE_STRING){
                    month = (int) row.getCell(37).getNumericCellValue();
                    year = (int) row.getCell(40).getNumericCellValue();
                } else {
                    month = (int) row.getCell(36).getNumericCellValue();
                    year = (int) row.getCell(39).getNumericCellValue();
                }

                row = rowIterator.next();
                //get full name from column 6
                String name = row.getCell(5).getStringCellValue();
                //split full name
                String forename = name.substring(name.indexOf(",") + 2);
                String surname = name.substring(0, name.indexOf(","));


                //skip irrelevant rows, select row with "Datum" in first cell
                while (row.getCell(0) == null || !row.getCell(0).getStringCellValue().equals("Datum")) {
                    row = rowIterator.next();
                }

                //store day outside of the loop the reuse it when first column is empty
                int day = 0;
                double netTimeWorked = 0.0;
                int forcedBreak = 0;

                MetaRecord metaRecord = new MetaRecord(0,"","",0,0.0,new ArrayList<>());

                //iterate through all records of the time tracking system for that employee
                while (rowIterator.hasNext()) {
                    //select first row with records
                    row = rowIterator.next();

                    cell = row.getCell(0);
                    /*when the first cell in a row contains the word "Monatssumme" that employees records have been
                    completed and the inner while loop is escaped
                    */
                    if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING &&
                            cell.getStringCellValue().equals("Monatssumme")) break;
                    //when there was no entry made for that day skip it (covers "Wochensumme" in first column)
                    if (row.getCell(8) == null) continue;
                    //rows containing F in column 15 can be ignored
                    if (row.getCell(14) != null && row.getCell(14).getStringCellValue().equals("F")) continue;

                    /*overwrite current day if necessary
                     Excel contains a format like (D)D,MM for the date, cast to int to cut of the month
                     */
                    day = cell == null ? day : (int) cell.getNumericCellValue();

                    //extract begin, end and pause; some values contain leading whitespace must be transformed to
                    // leading zero
                    String begin = row.getCell(8).getStringCellValue().replace(" ", "0").concat(":00");
                    String end = row.getCell(10).getStringCellValue().replace(" ", "0").concat(":00");
                    forcedBreak = row.getCell(12) == null ? forcedBreak :
                            Integer.parseInt(row.getCell(12).getStringCellValue().substring(3));

                    //extract netTime worked; column number is not consistent, check column for null value
                    // sometimes no net value is given at all (bug in hydra system)
                    if (cell != null && row.getCell(25) == null && row.getCell(26) == null) {
                        // TODO: 27.07.2017 post to logfile, maybe notification
                        netTimeWorked = 0;
                    } else {
                        netTimeWorked = cell == null ? netTimeWorked : row.getCell(26) == null ?
                                row.getCell(25).getNumericCellValue() :
                                row.getCell(26).getNumericCellValue();
                    }

                    if ( cell != null ){
                        metaRecord = new MetaRecord(id, forename, surname, forcedBreak, netTimeWorked, new ArrayList<>());
                        metaRecords.add(metaRecord);
                    }
                    metaRecord.getRecords().add(new Record(new DateTime(year+"-"+month+"-"+day+"T"+begin),
                            new DateTime(year+"-"+month+"-"+day+"T"+end)));
                }
            }
        } catch (Exception ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ein Fehler ist aufgetreten. \n" +
                    "Bitten wenden Sie sich an einen Administrator.\n\n Nachricht:\n " + ioe.getMessage());
        }
    }

    public static void main(String[] args) {

        //file dialog to choose file from system
        FileDialog fd = new FileDialog((java.awt.Frame) null, "Wählen Sie die Datei aus", FileDialog.LOAD);
        fd.setDirectory("C:\\");
        fd.setVisible(true);

        //get absolute file path from FileDialog
        String filepath = fd.getDirectory() + fd.getFile();
        //get file extension from file
        String fileExtension = fd.getFile().substring(fd.getFile().lastIndexOf(".") + 1);
        File inputFile = new File(filepath);

        //get output file name from user
        String filename = JOptionPane.showInputDialog("Geben Sie bitte den Namen der Ausgabe-Datei an.\n" +
                "Sonderzeichen werden entfernt.");
        //erase special characters from filename ( like .,;/\() )
        filename = filename.replaceAll("[^\\p{L}\\p{Z}]", "");
        File outputFile = new File(fd.getDirectory() + filename + ".csv");

        //check file extension for implementation
        if (fileExtension.equals("xlsx") || fileExtension.equals("xls")) {
            convert(inputFile, outputFile, fileExtension);
            JOptionPane.showMessageDialog(null, "Konvertierung war erfolgreich! \n\n" +
                    "Mitarbeiter gefunden: " + employeesFound + "\nBuchungen insgesamt: " + recordsFound);
        } else {
            JOptionPane.showMessageDialog(null, "Dateiformat wird derzeit nicht unterstützt.\n" +
                    "Bitte verwenden Sie nur '.xlsx' oder '.xls'");
        }
        System.exit(0);
    }
}