package cn.shepherd.excel.core.base;

import cn.shepherd.excel.annotation.Excel;
import cn.shepherd.excel.annotation.ExcelBigHead;
import cn.shepherd.excel.annotation.ExcelField;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.assertj.core.api.AssertionsForClassTypes.assertThat;

/**
 * @author Fuxin
 * @since 2019/11/23 16:02
 */

public class Tests {

    @Test
    void testExport() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        model.setTime(new Date());
        model.setExcelTime(DateUtil.convertTime("3:40:36"));
        data.add(model);
        data.add(model);
        data.add(model);
        exporterBase.appendSheet(Model.class, data);
        assertThat(workbook.getSheet("Summary")).isNotNull();
        OutputStream out = new FileOutputStream("target/foo.xls");
        workbook.write(out);
        out.close();
    }

    @Test
    @DisplayName("Same model class, different sheet name")
    void testExportMultiSheets() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        model.setTime(new Date());
        model.setExcelTime(DateUtil.convertTime("3:40:36"));
        data.add(model);
        data.add(model);
        data.add(model);
        exporterBase.appendSheet(Model.class, data)
                .appendSheet(Model.class, data, "Summary 2");
        assertThat(workbook.getSheet("Summary")).isNotNull();
        OutputStream out = new FileOutputStream("target/foo-multi.xls");
        workbook.write(out);
        out.close();
    }

    @Data
    @ExcelBigHead(value = "Just an example", fontStyle = ExporterBaseTest.CustomFontStyle.class)
    @Excel(value = "Summary", headerColor = 0x0D, herderHigh = 1024, rowHigh = 512, fontStyle = ExporterBaseTest.CustomFontStyle.class)
    public class Model {

        @ExcelField("Name")
        private String name;

        @ExcelField(value = "Birthday", type = ExcelField.CellType.DATE, width = 50, headerColor = 0x0C)
        private Date birthDate;

        @ExcelField(value = "Age", type = ExcelField.CellType.NUMERIC)
        private Integer age;

        @ExcelField(value = "Time", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private Date time;

        @ExcelField(value = "Excel Time", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private double excelTime;

    }
}
