package cn.shepherd.excel.core;

import cn.shepherd.excel.annotation.Excel;
import cn.shepherd.excel.annotation.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author Fuxin
 * @since 2019/11/23 13:48
 */
public class ExcelMetadata<T> {
    /**  */
    private final Class<T> clazz;

    /**  */
    private final Excel metaExcel;

    /**  */
    private final List<T> data;

    /**  */
    private final List<Field> excelFields;

    /**  */
    private final Workbook workbook;

    /**  */
    private final CreationHelper creationHelper;

    /**
     * 缓存每一列的样式，最多只能创建4000个cellStyle
     *
     * @see HSSFWorkbook#createCellStyle()
     **/
    private final Map<String, CellStyle> cellStyleMap;

    public ExcelMetadata(Class<T> clazz, List<T> data) {
        this(clazz, data, new HSSFWorkbook());
    }

    public ExcelMetadata(Class<T> clazz, List<T> data, Workbook workbook) {
        Excel excel = Objects.requireNonNull(clazz).getAnnotation(Excel.class);
        if (Objects.isNull(excel))
            throw new ExcelException(String.format("Can not get the @Excel annotation from this class %s", clazz.getName()));
        Objects.requireNonNull(data, "Data could not be null.");
        this.clazz = clazz;
        this.data = data;
        this.metaExcel = excel;
        this.excelFields = filterExcelFields(clazz);
        this.workbook = workbook;
        this.creationHelper = this.workbook.getCreationHelper();
        this.cellStyleMap = new HashMap<>();
    }

    public Class<T> getClazz() {
        return clazz;
    }

    public List<T> getData() {
        return data;
    }

    public Excel getMetaExcel() {
        return metaExcel;
    }

    public List<Field> getExcelFields() {
        return excelFields;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public CreationHelper getCreationHelper() {
        return creationHelper;
    }

    public Map<String, CellStyle> getCellStyleMap() {
        return cellStyleMap;
    }

    private List<Field> filterExcelFields(Class<T> clazz) {
        return Stream.of(clazz.getDeclaredFields())
                .peek(field -> field.setAccessible(true))
                .filter(field -> field.isAnnotationPresent(ExcelField.class))
                .collect(Collectors.toList());
    }
}
