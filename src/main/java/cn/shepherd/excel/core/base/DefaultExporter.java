package cn.shepherd.excel.core.base;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Fuxin
 * @since 2019/11/23 16:58
 */
public class DefaultExporter extends ExporterBase {

    public DefaultExporter(Workbook workbook) {
        super(workbook);
    }
}
