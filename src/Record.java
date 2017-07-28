import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

/**
 * This class inherits all necessary attributes to reconstruct the corresponding time tracking system and
 * convert it to the given format. Contains data like employee, break and net time worked.
 * The detailed beginning and ending times are store in a list of intervals.
 */
@Data
@AllArgsConstructor
class Record {
    private int id; // == Transpondernummer
    private String forename;
    private String surname;
    private int forcedBreak;
    private double netTimeWorked;
    private List<Interval> intervals;
}
