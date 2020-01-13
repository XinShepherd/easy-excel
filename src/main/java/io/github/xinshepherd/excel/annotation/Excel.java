package io.github.xinshepherd.excel.annotation;

import io.github.xinshepherd.excel.core.DefaultFontStyle;
import io.github.xinshepherd.excel.core.FontStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author Fuxin
 * @since 2019/11/23 10:48
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {

    /**
     * 设置sheet名称
     *
     * @return sheet名称
     */
    String value() default "";

    /**
     * 全局设置表头单元格颜色
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
     * 设置所有数据行的高度, 默认 255
     *
     * @see Sheet#getDefaultRowHeight()
     * @see org.apache.poi.ss.usermodel.Row#setHeight
     * @return 数据行高度
     */
    short rowHigh() default -1;

    /**
     * 设置头部行的高度，默认 255
     *
     * @see Sheet#getDefaultRowHeight()
     * @see org.apache.poi.ss.usermodel.Row#setHeight
     * @return 头部行高度
     */
    short herderHigh() default -1;

    /**
     * 设置头部行的字体样式
     *
     * @see FontStyle
     * @return 头部字体样式
     */
    Class<? extends FontStyle> fontStyle() default DefaultFontStyle.class;


    /**
     * 设置是否冻结表头
     *
     * @return true or false
     */
    boolean freezePane() default true;
}
