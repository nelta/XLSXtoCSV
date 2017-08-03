import org.joda.time.DateTime;
import org.joda.time.LocalTime;
import org.joda.time.Period;

import java.util.ArrayList;
import java.util.List;


public class Testarea {
    public static void main(String[] args) {
        String a = "13:56:00";
        String b = "00:25:00";
        String MIDNIGHT = "00:00:00";

        LocalTime test = new LocalTime(a);
        DateTime aDate = new DateTime("2017-08-01T"+a);
        DateTime bDate = new DateTime("2017-08-01T"+b);
        System.out.println(bDate.minusMinutes(30));
    }
}
