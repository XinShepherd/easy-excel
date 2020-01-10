package cn.shepherd.excel.annotation;

import cn.shepherd.excel.core.DefaultFontStyle;
import cn.shepherd.excel.core.FontStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author Fuxin
 * @since 2019/11/23 10:16
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    /**
     * 设置列名
     *
     * @return 列名
     */
    String value() default "";

    /**
     * 设置单元格类型，默认文本类型
     *
     * @return 单元格类型
     */
    CellType type() default CellType.TEXT;

    /**
     * 设置日期类型格式，单元格类型为日期类型时有效
     *
     * @return 日期类型格式
     */
    String datePattern() default "yyyy-MM-dd HH:mm:ss";

    /**
     * 设置列宽度，单位(0.5 个字符)
     * eg: 当设置为10时，相当于该列为5个字符的宽度
     *
     * @return 列宽度
     */
    int width() default 20;

    /**
     * 设置单元格垂直对齐方式
     *
     * @return 垂直对齐方式，默认居中
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

    /**
     * 设置单元格水平对齐方式
     *
     * @return 水平对齐方式，默认居中
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.CENTER;

    /**
     * 设置表头单元格颜色
     * 请参考 {@link HSSFColor#getIndex()} 以及 {@link HSSFColor.HSSFColorPredefined#getIndex()}
     *
     * 如果需要自定义颜色，请使用调色板 {@link org.apache.poi.hssf.usermodel.HSSFPalette}
     * 例如：
     * <code>
     *   //creating a custom palette for the workbook
     *   HSSFPalette palette = wb.getCustomPalette();
     *   palette.setColorAtIndex(index, r, b, g);
     *   cellStyle.setFillForegroundColor(index);
     * </code>
     *
     * @return 颜色
     */
    short headerColor() default -1;

    /**
     * 设置数据单元格颜色
     * 请参考 {@link HSSFColor#getIndex()} 以及 {@link HSSFColor.HSSFColorPredefined#getIndex()}
     *
     * 如果需要自定义颜色，请使用调色板 {@link org.apache.poi.hssf.usermodel.HSSFPalette}
     * 例如：
     * <code>
     *   //creating a custom palette for the workbook
     *   HSSFPalette palette = wb.getCustomPalette();
     *   palette.setColorAtIndex(index, r, b, g);
     *   cellStyle.setFillForegroundColor(index);
     * </code>
     *
     * @return 数据单元格颜色
     */
    short color() default -1;

    /**
     * 设置单元格字体
     *
     * @see FontStyle
     * @return 单元格字体
     */
    Class<? extends FontStyle> fontStyle() default DefaultFontStyle.class;

    enum CellType {
        TEXT,
        NUMERIC,
        DATE,
        TIME // 不包含日期
    }
}
