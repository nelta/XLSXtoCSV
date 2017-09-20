import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.Period;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.nio.charset.Charset;
import java.util.*;

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
    private static Set<Record> recordsToBeReviewed = new HashSet<>();

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
                    "Mitarbeiter gefunden: " + employeesFound + "\nEinträge insgesamt: " + recordsFound + "\n\n" +
                    "Es gibt " + recordsToBeReviewed.size() + " Einträge, die manuell angepasst werden müssen.\n" +
                    "Nähere Informationen im Bericht");
        } else {
            JOptionPane.showMessageDialog(null, "Dateiformat wird derzeit nicht unterstützt.\n" +
                    "Bitte verwenden Sie nur '.xlsx' oder '.xls'");
        }
        System.exit(0);
    }

    private static void convert(File inputFile, File outputFile, String fileExtension) {

        //extract periodes
        extractRecords(inputFile, fileExtension);

        //handle exceptions, create logfile
        modifyRecords();
        createLogfile();

        //convert periods and write to new file
        createOutput(outputFile);

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

            //Iterator to iterate through each row of the sheet
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
                    if (row.getCell(8) == null || row.getCell(8).getStringCellValue().equals("")) continue;
                    //rows containing F in column 15 can be ignored
                    if (row.getCell(14) != null && row.getCell(14).getStringCellValue().equals("F")) continue;

                    /*overwrite current day if necessary
                     Excel contains a format like (D)D,MM for the date, cast to int to cut of the month
                     */
                    day = cell == null || cell.getNumericCellValue() == 0 ? day : (int) cell.getNumericCellValue();

                    //extract begin, end and pause; some values contain leading whitespace must be transformed to
                    // leading zero
                    String begin = row.getCell(8).getStringCellValue().replace(" ", "0").concat(":00");
                    String end = row.getCell(10).getStringCellValue().replace(" ", "0").concat(":00");
                    forcedBreak = row.getCell(12) == null || row.getCell(12).getStringCellValue().equals("") ?
                            forcedBreak : Integer.parseInt(row.getCell(12).getStringCellValue().substring(3));

                    //extract netTime worked; column number is not consistent, check column for null value
                    // sometimes no net value is given at all (bug in hydra system)
                    if (cell != null && ( (row.getCell(25) == null && row.getCell(26) == null) ||
                            (row.getCell(25).getCellType() == 3 && row.getCell(26).getCellType() == 3))
                            ) {
                        netTimeWorked = 0;
                    } else {
                        netTimeWorked = cell == null || cell.getNumericCellValue()==0 ? netTimeWorked :
                                row.getCell(26) == null || row.getCell(25).getCellType()==0 ?
                                        row.getCell(25).getNumericCellValue() :
                                        row.getCell(26).getNumericCellValue();
                    }

                    //create a new record if the first column has a new date (all other possibilities for nonnull values
                    //in the first column should have been covered and escaped before)
                    if (cell != null && cell.getNumericCellValue() != 0) {
                        record = new Record(id, forename, surname, forcedBreak, netTimeWorked, new ArrayList<>());
                        records.add(record);
                        recordsFound++;
                    }

                    //System.out.println(year + "-" + month + "-" + day + "T" + begin);
                    DateTime beginDate = new DateTime(year + "-" + month + "-" + day + "T" + begin);
                    DateTime endDate = new DateTime(year + "-" + month + "-" + day + "T" + end);

                    //add new period to the record
                    record.getIntervals().add(new Interval(beginDate, endDate));
                }
                employeesFound++;
            }
        } catch (Exception ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ein Fehler ist aufgetreten. \n" +
                    "Bitten wenden Sie sich an einen Administrator.\n\n Nachricht:\n " + ioe.getMessage());
        }
    }

    private static void modifyRecords() {
        final int MINUTES_OF_DAY = 1440;
        for (Record r : records) {
            ArrayList<Interval> intervals = r.getIntervals();

            /*
            Adjusting night-shift:
            its a night-shift, if the end is before the start (dates have not been modified yet)
            and if the interval before the change of days is less than afterwards
            if it is a night shift all DateTimes before the change of days have to be reduced by 1 due to an "error"
            at Stryker's time tracking system: if an employees starts at 22:34, 01.08. and is working till 6:34, 02.08.
            all entries are stored under the 2nd of August
            therefor for all records that include a day change but are no night-shifts have to be modified
             */
            final DateTime begin = intervals.get(0).getBegin();
            final DateTime end = intervals.get(intervals.size() - 1).getEnd();
            if (begin.compareTo(end) > 0) {
                if ((MINUTES_OF_DAY - begin.getMinuteOfDay()) < end.getMinuteOfDay()) {
                    for (Interval i : intervals) {
                        i.setBegin(i.getBegin().compareTo(begin) >= 0 ? i.getBegin().minusDays(1) : i.getBegin());
                        i.setEnd(i.getEnd().compareTo(begin) > 0 ? i.getEnd().minusDays(1) : i.getEnd());
                    }
                } else {
                    for (Interval i : intervals) {
                        i.setBegin(i.getBegin().compareTo(end) < 0 ? i.getBegin().plusDays(1) : i.getBegin());
                        i.setEnd(i.getEnd().compareTo(end) <= 0 ? i.getEnd().plusDays(1) : i.getEnd());
                    }
                }
            }

            //if the employee doesn't have the permission to start before his shift begins, those minutes don't count
            double calculatedNetWorkingTime = calculateNetWorkingTime(r);
            int minutesBelowCalculations = (int) Math.round((calculatedNetWorkingTime - r.getNetTimeWorked()) * 60);

            if (r.getNetTimeWorked() < calculatedNetWorkingTime) {
                DateTime beginRecord = intervals.get(0).getBegin();
                if (minutesBelowCalculations + beginRecord.getMinuteOfHour() == 60) {
                    beginRecord = beginRecord.plusHours(1).minusMinutes(beginRecord.getMinuteOfHour());
                    intervals.get(0).setBegin(beginRecord);
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
            }

            if (r.getForcedBreak() > 0) {
                boolean breakEntered = false;
                for (Interval interval : intervals) {
                    Period period = new Period(interval.getBegin(), interval.getEnd());
                    if (period.getHours() * 60 + period.getMinutes() > r.getForcedBreak()) {
                        interval.setEnd(interval.getEnd().minusMinutes(r.getForcedBreak()));
                        r.setForcedBreak(0);
                        breakEntered = true;
                        break;
                    }
                }
                if (!breakEntered) {
                    recordsToBeReviewed.add(r);
                }
            }
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
                    .append(" Inkonsistenzen bei den Einträgen gefunden, die eine manuelle Prüfung erfordern: \r\n\r\n");

            ArrayList<Record> sortedRecords = new ArrayList<>(recordsToBeReviewed);
            Collections.sort(sortedRecords);

            int count = 1;
            for (Record r : sortedRecords) {
                data.append(count).append("\t");
                count++;
                double calculatedNetWorkingTime = calculateNetWorkingTime(r);
                double dif = calculatedNetWorkingTime - r.getNetTimeWorked();
                data.append("Beim Mitarbeiter: ")
                        .append(r.getId())
                        .append("\t")
                        .append(r.getForename())
                        .append(" ")
                        .append(r.getSurname())
                        .append("\r\n");
                data.append("\tBeim Eintrag am: ")
                        .append(r.getIntervals().get(0).getBegin().getDayOfMonth())
                        .append(".")
                        .append(r.getIntervals().get(0).getBegin().getMonthOfYear())
                        .append(".")
                        .append(r.getIntervals().get(0).getBegin().getYearOfCentury())
                        .append("\r\n");
                data.append("\tDifferenz in Minuten: ")
                        .append(Math.round((calculatedNetWorkingTime - r.getNetTimeWorked()) * 60))
                        .append("min / als Zahl: ")
                        .append(Math.round(dif * 100) / 100.0)
                        .append("\r\n");

                //append known issues
                if (r.getNetTimeWorked() == 0) {
                    data.append("\tProblem: KEINE IST-ZEIT IN TABELLE EINGETRAGEN!")
                            .append("\r\n");
                }

                if (r.getForcedBreak() > 0){
                    data.append("\tProblem: PAUSE KONNTE NICHT EINGETRAGEN WERDEN!")
                            .append("\r\n");
                }

                data.append("\r\n");
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

    private static void createOutput(File outputFile) {
        final String KOMMEN = "K";
        final String GEHEN = "G";
        final DateTimeFormatter format = DateTimeFormat.forPattern("dd.MM.yyyy HH:mm:ss");

        //Stringbuilder for storing data into CSV files
        StringBuilder data = new StringBuilder();
        try {
            Writer fos = new OutputStreamWriter(new FileOutputStream(outputFile), ASCII);

            for (Record r : records) {
                for (Interval i : r.getIntervals()) {
                    data.append(KOMMEN)
                            .append(";")
                            .append(r.getId())
                            .append(";")
                            .append(format.print(i.getBegin()))
                            .append(";;;;;\r\n");
                    data.append(GEHEN)
                            .append(";")
                            .append(r.getId())
                            .append(";")
                            .append(format.print(i.getEnd()))
                            .append(";;;;;\r\n");
                }
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

}