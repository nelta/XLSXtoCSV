import lombok.AllArgsConstructor;
import lombok.Data;
import org.joda.time.DateTime;


/**
 * This class inherits all necessary attributes to reconstruct the corresponding time tracking system and
 * convert it to the given format
 */
@Data
@AllArgsConstructor
class Interval {
    private DateTime begin;
    private DateTime end;
}
