package io.github.xinshepherd.excel.core.base;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Fuxin
 * @since 1.1.0
 */
public class DefaultExporter extends ExporterBase {

    public DefaultExporter(Workbook workbook) {
        super(workbook);
    }
}
