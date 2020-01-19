package io.github.xinshepherd.excel;

import io.github.xinshepherd.excel.annotation.Excel;
import io.github.xinshepherd.excel.annotation.ExcelBigHead;
import io.github.xinshepherd.excel.annotation.ExcelField;
import io.github.xinshepherd.excel.core.CellStyleProcessor;
import io.github.xinshepherd.excel.core.base.DefaultExporter;
import io.github.xinshepherd.excel.core.base.ExporterBase;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;

import static org.apache.poi.ss.usermodel.IndexedColors.ROSE;
import static org.assertj.core.api.AssertionsForClassTypes.assertThat;

/**
 * @author Fuxin
 * @since 2019/11/23 16:02
 */

public class Tests {

    @Test
    void testExport() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        model.setTime(new Date());
        model.setExcelTime(DateUtil.convertTime("3:40:36"));
        model.setStringTime("3:40:36");
        model.setJavaTime(LocalTime.now());
        data.add(model);
        data.add(model);
        data.add(model);
        exporterBase.appendSheet(Model.class, data);
        assertThat(workbook.getSheet("Summary")).isNotNull();
        OutputStream out = new FileOutputStream("target/foo.xlsx");
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
        model.setAge(9);
        model.setTime(new Date());
        model.setExcelTime(DateUtil.convertTime("3:40:36"));
        model.setStringTime("3:40:36");
        model.setJavaTime(LocalTime.now());
        data.add(model);
        Model model2 = new Model();
        model2.setName("foo");
        model2.setBirthDate(new Date());
        model2.setAge(20);
        data.add(model2);
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
    @Excel(value = "Summary",
            headerColor = 0x0D,
            herderHigh = 1024,
            rowHigh = 512,
            fontStyle = ExporterBaseTest.CustomFontStyle.class,
            colSplit = 2,
            rowSplit = 2)
    public class Model {

        @ExcelField("Name")
        private String name;

        @ExcelField(value = "Birthday", type = ExcelField.CellType.DATE, width = 50, headerColor = 0x0C)
        private Date birthDate;

        @ExcelField(value = "Age", type = ExcelField.CellType.NUMERIC, customStyle = CustomCellStyleProcessor.class)
        private Integer age;

        @ExcelField(value = "Time", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private Date time;

        @ExcelField(value = "Excel Time", type = ExcelField.CellType.TIME, datePattern = "h:mm:ss")
        private double excelTime;

        @ExcelField(value = "String Time", type = ExcelField.CellType.TIME, datePattern = "h:mm:ss")
        private String stringTime;

        @ExcelField(value = "Java Time", type = ExcelField.CellType.TIME, datePattern = "h:mm:ss")
        private LocalTime javaTime;

    }

    public static class CustomCellStyleProcessor implements CellStyleProcessor {

        private static final String NONE = "NONE";
        private static final String ADULT = "ADULT";
        private static final String CHILD = "CHILD";

        @Override
        public String getLabel(Object data) {
            if (Objects.isNull(data))
                return NONE;
            Integer age = (Integer) data;
            return age < 18 ? CHILD : ADULT;
        }

        @Override
        public CellStyle customize(CellStyle cellStyle, String label) {
            if (ADULT.equals(label)) {
                cellStyle.setFillForegroundColor(ROSE.index);
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            return cellStyle;
        }
    }

}
