import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.Period;

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
    private static String filename;
    private static FileDialog fd;

    //charset required by Landwehr
    private final static Charset ASCII = Charset.forName("ASCII");
    //charset for logfile
    private final static Charset UTF8 = Charset.forName("UTF8");

    //List of extracted periodes
    private static ArrayList<Record> records = new ArrayList<>();
    private static ArrayList<Record> recordsToBeReviewed = new ArrayList<>();

    private static void convert(File inputFile, File outputFile, String fileExtension) {

        //extract periodes
        extractRecords(inputFile, fileExtension);

        //handle exceptions, create logfile
        handleExceptions();
        createLogfile();

        //convert periods and write to new file
        // TODO: 27.07.2017 post to logfile, maybe notification
        // TODO: escape 0 netTime in filewriting and calculations

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

                //define the current date, used when multiple periodes have been made for one day
                String curDate = "";

                while (rowIterator.hasNext()) {
                    //select first row with periodes
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

                    //count total periodes
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

    private static void createLogfile() {
        StringBuilder data = new StringBuilder();
        File logFile = new File(fd.getDirectory() + "LOG-" + filename + ".txt");
        try {
            Writer fos = new OutputStreamWriter(new FileOutputStream(logFile), UTF8);

            if (recordsToBeReviewed.size() == 0) {
                data.append("Es wurden keine Inkonsistenzen gefunden.");
                fos.write(data.toString());
                fos.close();
                return;
            }

            data.append("Folgende ")
                    .append(recordsToBeReviewed.size())
                    .append(" Inkonsistenzen, die eine manuelle Prüfung erfordern, wurden gefunden: \r\n\r\n");

            int count = 1;
            for (Record r : recordsToBeReviewed) {
                data.append(count).append("\t");
                count++;
                double calculatedNetWorkingTime = calculateNetWorkingTime(r);
                double dif = calculatedNetWorkingTime - r.getNetTimeWorked();
                data.append("Differenz in Minuten / als Zahl: ")
                        .append(Math.round((calculatedNetWorkingTime - r.getNetTimeWorked()) * 60))
                        .append("min / ")
                        .append(dif)
                        .append("\r\n");
                data.append("Bei Eintrag: ")
                        .append(r)
                        .append("\r\n\r\n");
            }

            //write everything to logfile
            fos.write(data.toString());
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ein Fehler ist aufgetreten. \n" +
                    "Bitten wenden Sie sich an einen Administrator.\n\n Nachricht:\n " + e.getMessage());
        }
    }

    private static void handleExceptions() {
        for (Record r : records) {

            //if the employee doesn't have the permission to start before his shift begins, those minutes doesn't count
            double calculatedNetWorkingTime = calculateNetWorkingTime(r);
            int minutesBelowCalculations = (int) Math.round((calculatedNetWorkingTime - r.getNetTimeWorked()) * 60);
            if (r.getNetTimeWorked() < calculatedNetWorkingTime) {
                DateTime beginRecord = r.getIntervals().get(0).getBegin();
                if (minutesBelowCalculations + beginRecord.getMinuteOfHour() == 60) {
                    beginRecord = beginRecord.plusHours(1).minusMinutes(beginRecord.getMinuteOfHour());
                    r.getIntervals().get(0).setBegin(beginRecord);
                }
            }

            //handle if the given break is ignored by time calculations
            if (minutesBelowCalculations * (-1) == r.getForcedBreak()) {
                r.setForcedBreak(0);
            }

            //recalculate net working time
            calculatedNetWorkingTime = calculateNetWorkingTime(r);

            if (calculatedNetWorkingTime != r.getNetTimeWorked()) {
                recordsToBeReviewed.add(r);
                System.out.println(r.getNetTimeWorked() + " : " + calculatedNetWorkingTime + " - Minuten: " +
                        minutesBelowCalculations + " Record: " + r);
            }
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
                if (row.getCell(39) != null && row.getCell(39).getCellType() == Cell.CELL_TYPE_STRING) {
                    month = (int) row.getCell(38).getNumericCellValue();
                    year = (int) row.getCell(41).getNumericCellValue();
                } else if (row.getCell(38) != null && row.getCell(38).getCellType() == Cell.CELL_TYPE_STRING) {
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

                Record record = new Record(0, "", "", 0, 0.0, new ArrayList<>());

                //iterate through all periodes of the time tracking system for that employee
                while (rowIterator.hasNext()) {
                    //select first row with periodes
                    row = rowIterator.next();

                    cell = row.getCell(0);
                    /*when the first cell in a row contains the word "Monatssumme" that employees periodes have been
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
                        netTimeWorked = 0;
                    } else {
                        netTimeWorked = cell == null ? netTimeWorked : row.getCell(26) == null ?
                                row.getCell(25).getNumericCellValue() :
                                row.getCell(26).getNumericCellValue();
                    }

                    //create a new record if the first column has a new date (all other possibilities for nonnull values
                    //in the first column should have been covered and escaped before)
                    if (cell != null) {
                        record = new Record(id, forename, surname, forcedBreak, netTimeWorked, new ArrayList<>());
                        records.add(record);
                    }

                    DateTime beginDate = new DateTime(year + "-" + month + "-" + day + "T" + begin);
                    DateTime endDate = new DateTime(year + "-" + month + "-" + day + "T" + end);

                    if (beginDate.getHourOfDay() > endDate.getHourOfDay()) endDate = endDate.plusDays(1);

                    //add new period to the record
                    record.getIntervals().add(new Interval(beginDate, endDate));
                }
            }
        } catch (Exception ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ein Fehler ist aufgetreten. \n" +
                    "Bitten wenden Sie sich an einen Administrator.\n\n Nachricht:\n " + ioe.getMessage());
        }
    }

    private static double calculateNetWorkingTime(Record record) {
        double calculatedNetWorkingTime;
        Period p = new Period();
        for (Interval i : record.getIntervals()) {
            final Period temp = new Period(i.getBegin(), i.getEnd());
            p = p.plusHours(temp.getHours()).plusMinutes(temp.getMinutes());
        }
        while (p.getMinutes() >= 60) p = p.plusHours(1).minusMinutes(60);
        p = p.getMinutes() >= record.getForcedBreak() ?
                p.minusMinutes(record.getForcedBreak()) :
                p.plusMinutes(60 - record.getForcedBreak()).minusHours(1);
        calculatedNetWorkingTime = p.getHours() + p.getMinutes() / 60.0;
        return Math.round(calculatedNetWorkingTime * 100) / 100.0;
    }

    public static void main(String[] args) {

        String fileExtension;
        File inputFile;
        //catch NullPointerException that is thrown when the User aborts or closes the dialog
        try {
            //file dialog to choose file from system
            fd = new FileDialog((Frame) null, "Wählen Sie die Datei aus", FileDialog.LOAD);
            fd.setDirectory("C:\\");
            fd.setVisible(true);

            //get absolute file path from FileDialog
            String filepath = fd.getDirectory() + fd.getFile();
            //get file extension from file
            fileExtension = fd.getFile().substring(fd.getFile().lastIndexOf(".") + 1);
            inputFile = new File(filepath);

        } catch (NullPointerException ne) {
            System.out.println("---Programmabbruch---");
            System.exit(0);
            return;
        }
        //get output file name from user
        filename = JOptionPane.showInputDialog("Geben Sie bitte den Namen der Ausgabe-Datei an.\n" +
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