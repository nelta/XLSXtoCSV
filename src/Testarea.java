import org.joda.time.DateTime;
import org.joda.time.LocalTime;
import org.joda.time.Period;

import java.util.ArrayList;
import java.util.List;


public class Testarea {
    public static void main(String[] args) {
        String a = "13:56:00";
        String b = "05:25:00";
        String MIDNIGHT = "00:00:00";

        LocalTime test = new LocalTime(a);
        DateTime aDate = new DateTime("2017-08-01T"+test);
        DateTime bDate = new DateTime("2017-08-01T"+b);
        if ( aDate.getHourOfDay() > bDate.getHourOfDay())
            bDate = bDate.plusDays(1);
        System.out.println(new Period().plusDays(1).plusHours(1).plusMinutes(5));
        System.out.println(bDate);
        System.out.println(new Period(aDate, bDate));
        System.out.println(a);

        System.out.println((int) 4.5);

        String da = " 0:30";
        System.out.println(da.replace(" ", "0"));
        List<Record> records = new ArrayList<>();
        Record record = new Record(0, "", "dsf", 0, 0.0, new ArrayList<>());
        records.add(record);
        record = new Record(1, "sads", "", 0, 0.0, new ArrayList<>());
        records.add(record);
        System.out.println(records);
    }
}
