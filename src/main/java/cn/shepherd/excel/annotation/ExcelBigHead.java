package cn.shepherd.excel.annotation;

import cn.shepherd.excel.core.DefaultFontStyle;
import cn.shepherd.excel.core.FontStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author Fuxin
 * @since 2020/1/10 17:01
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelBigHead {

    String value();

    /**
     * 设置头部行的字体样式
     *
     * @see FontStyle
     * @return 头部字体样式
     */
    Class<? extends FontStyle> fontStyle() default DefaultFontStyle.class;

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

    int fromRow() default 0;

    int toRow() default 0;

    int fromColumn() default 0;

    int toColumn() default 1;

    /**
     * 列数是否自动对齐表头
     *
     * @return 是则不会根据 fromColumn 和 toColumn 设置合并的单元格列数
     * @see #fromColumn()
     * @see #toColumn()
     */
    boolean autoMergeColumn() default true;

}
