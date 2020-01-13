package cn.shepherd.excel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @author Fuxin
 * @since 2020/1/13 11:29
 */
public class ExcelSheetBuilder<T> {

    private Class<T> modelClass;

    private List<T> data;

    private Workbook workbook;

    private String sheetName;

    public ExcelSheetBuilder(Class<T> modelClass, List<T> data, Workbook workbook) {
        this.modelClass = modelClass;
        this.data = data;
        this.workbook = workbook;
    }

    public ExcelSheetBuilder<T> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelSheetMetadata<T> build() {
        return new ExcelSheetMetadata<>(modelClass, data, workbook, sheetName);
    }
}
