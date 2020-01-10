package cn.shepherd.excel.core.util;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.LocaleUtil;

import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.temporal.ChronoField;
import java.time.temporal.TemporalAccessor;
import java.time.temporal.TemporalQueries;
import java.util.Calendar;
import java.util.Date;

/**
 * @author Fuxin
 * @since 2019/11/23 19:44
 */
public abstract class DateTimeUtil extends DateUtil {

    public static final DateTimeFormatter TIME_FORMATTER = DateTimeFormatter.ofPattern("HH:mm:ss");

    private static final DateTimeFormatter dateTimeFormats = new DateTimeFormatterBuilder()
            .appendPattern("[[yyyy-]MM-dd][[ ]h:m[:s] a][[ ][H]H:m[m][:s[s]]]")
            .appendPattern("[dd MMM[ yyyy]][[ ]h:m[:s] a][[ ]H:m[:s]]")
            .appendPattern("[[yyyy ]dd-MMM[-yyyy]][[ ]h:m[:s] a][[ ]H:m[:s]]")
            .appendPattern("[M/dd[/yyyy]][[ ]h:m[:s] a][[ ]H:m[:s]]")
            .appendPattern("[[yyyy/]M/dd][[ ]h:m[:s] a][[ ]H:m[:s]]")
            .parseDefaulting(ChronoField.YEAR_OF_ERA, LocaleUtil.getLocaleCalendar().get(Calendar.YEAR))
            .toFormatter();


    /**
     * Converts a temporal to its (Excel) numeric equivalent
     *
     * @see DateUtil#convertTime(String)
     *
     * @param temporal the temporal object to format, not null
     * @return a double between 0 and 1 representing the fraction of the day
     *
     * @since 1.0.0
     */
    public static double convertTime(TemporalAccessor temporal) {
        return convertTime(TIME_FORMATTER.format(temporal));
    }

    public static Double parseDateTime(String str){
        TemporalAccessor tmp = dateTimeFormats.parse(str.replaceAll("\\s+", " "));
        LocalTime time = tmp.query(TemporalQueries.localTime());
        LocalDate date = tmp.query(TemporalQueries.localDate());
        if(time == null && date == null) return null;

        double tm = 0;
        if(date != null) {
            Date d = Date.from(date.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());
            tm = DateUtil.getExcelDate(d);
        }
        if(time != null) tm += 1.0*time.toSecondOfDay()/SECONDS_PER_DAY;

        return tm;
    }

}
