package cn.shepherd.excel.core.base;

import cn.shepherd.excel.core.ExcelSheetBuilder;
import cn.shepherd.excel.core.ExcelSheetMetadata;
import cn.shepherd.excel.core.SheetCreator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @author Fuxin
 * @since 2019/11/23 10:44
 */
public abstract class ExporterBase {

    private final Workbook workbook;

    public ExporterBase(Workbook workbook) {
        this.workbook = workbook;
    }

    public <T> ExporterBase appendSheet(Class<T> clazz, List<T> data) {
        return this.appendSheet(clazz, data, null);
    }

    public <T> ExporterBase appendSheet(Class<T> clazz, List<T> data, String sheetName) {
        ExcelSheetMetadata<T> metadata = new ExcelSheetBuilder<>(clazz, data, workbook)
                .sheetName(sheetName)
                .build();
        SheetCreator<T> sheetCreator = new SheetCreator<>(metadata);
        sheetCreator.createSheet();
        return this;
    }

}
