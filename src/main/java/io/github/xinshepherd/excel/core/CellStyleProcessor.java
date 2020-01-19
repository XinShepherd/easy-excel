package io.github.xinshepherd.excel.core;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 实现类的构造器不能带有参数，否则会抛错
 *
 * @see ExcelException
 * @author Fuxin
 * @since 1.2.0
 */
public interface CellStyleProcessor {

    /**
     * 设置 CellStyle 的标签，主要是为了缓存 CellStyle 而设计
     *
     * @param data 单元格数据
     * @return 标签
     */
    String getLabel(Object data);

    /**
     * 自定义样式
     * @param cellStyle 样式
     * @param label 标签 {@link #getLabel(Object)}
     * @return 自定义的样式
     */
    CellStyle customize(CellStyle cellStyle, String label);

}
