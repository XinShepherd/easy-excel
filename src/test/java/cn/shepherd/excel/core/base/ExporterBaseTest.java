package cn.shepherd.excel.core.base;

import cn.shepherd.excel.annotation.Excel;
import cn.shepherd.excel.annotation.ExcelField;
import cn.shepherd.excel.core.ExcelSheetMetadata;
import cn.shepherd.excel.core.FontStyle;
import cn.shepherd.excel.core.util.DateTimeUtil;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

import static org.assertj.core.api.AssertionsForClassTypes.assertThat;


/**
 * @author Fuxin
 * @since 2019/11/23 11:10
 */
class ExporterBaseTest {

    @Test
    void testGenerateEmptyWorkBook() {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        exporterBase.appendSheet(Model.class, data);
        assertThat(workbook.getSheetAt(0)).isNotNull();
    }

    @Test
    void testGenerateWorkBookWithData() {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        data.add(model);
        data.add(model);
        data.add(model);
        exporterBase.appendSheet(Model.class, data);
        assertThat(workbook.getSheetAt(0)).isNotNull();
        assertThat(workbook.getSheet("汇总表")).isNotNull();
        assertThat(workbook.getSheet("汇总表").getLastRowNum()).isEqualTo(3);
    }

    @Test
    void testGenerateWorkBookWithDataIntoFile() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        model.setTime(new Date(3600000));
        model.setExcelTime(DateUtil.convertTime("3:40:36"));
        data.add(model);
        data.add(model);
        data.add(model);
        exporterBase.appendSheet(Model.class, data);
        assertThat(workbook.getSheet("汇总表")).isNotNull();
        // 定义输出流
        OutputStream out = new FileOutputStream("target/foo1.xls");
        workbook.write(out);
        out.close();
    }

    @Test
    void testAppendWorkBookWithDataIntoFile() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        model.setTime(new Date(3600000));
        model.setExcelTime(DateUtil.convertTime("3:40:36"));
        data.add(model);
        data.add(model);
        data.add(model);
        exporterBase.appendSheet(Model.class, data);
        assertThat(workbook.getSheet("汇总表")).isNotNull();
        List<Detail> detailList = new ArrayList<>();
        Detail detail = new Detail();
        detail.setName("foo");
        detail.setBirthDate(new Date());
        detail.setAge(10);
        detail.setTime(LocalDateTime.now());
        detail.setExcelTime(DateUtil.convertTime("3:45:36"));
        detailList.add(detail);
        detailList.add(detail);
        detailList.add(detail);
        exporterBase.appendSheet(Detail.class, detailList);
        assertThat(workbook.getSheet("详情表")).isNotNull();
        // 定义输出流
        OutputStream out = new FileOutputStream("target/foo1.xls");
        workbook.write(out);
        out.close();
    }

    @Test
    void testGenerateWorkBook() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        ExporterBase exporterBase = new DefaultExporter(workbook);
        List<Model> data = new ArrayList<>();
        Model model = new Model();
        model.setName("foo");
        model.setBirthDate(new Date());
        model.setAge(10);
        model.setTime(new Date(3600000));
        model.setExcelTime(DateTimeUtil.convertTime("3:40:36"));
        data.add(model);
        data.add(model);
        data.add(model);
        List<Detail> detailList = new ArrayList<>();
        Detail detail = new Detail();
        detail.setName("foo");
        detail.setBirthDate(new Date());
        detail.setAge(10);
        detail.setTime(LocalDateTime.now());
        detail.setExcelTime(DateTimeUtil.convertTime("12:45:36"));
        detail.setDateStr(LocalDate.now().toString());
        detailList.add(detail);
        detailList.add(detail);
        detailList.add(detail);
        Map<Class, List> dataMap = new LinkedHashMap<>();
        exporterBase.appendSheet(Model.class, data)
                .appendSheet(Detail.class, detailList);
        // 定义输出流
        OutputStream out = new FileOutputStream("target/foo2.xls");
        workbook.write(out);
        out.close();
    }

    @Data
    @Excel(value = "汇总表", headerColor = 0x0D, herderHigh = 1024, rowHigh = 512, fontStyle = CustomFontStyle.class)
    class Model {

        @ExcelField("姓名")
        private String name;

        @ExcelField(value = "出生日期", type = ExcelField.CellType.DATE, width = 50, headerColor = 0x0C)
        private Date birthDate;

        @ExcelField(value = "年龄", type = ExcelField.CellType.NUMERIC)
        private Integer age;

        @ExcelField(value = "时间", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private Date time;

        @ExcelField(value = "excel时间", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private double excelTime;

    }

    @Data
    @Excel(value = "详情表", headerColor = 0x0D, herderHigh = 1024, rowHigh = 512)
    class Detail {

        @ExcelField("姓名")
        private String name;

        @ExcelField(value = "出生日期", type = ExcelField.CellType.DATE, width = 50, headerColor = 0x0C)
        private Date birthDate;

        @ExcelField(value = "年龄", type = ExcelField.CellType.NUMERIC)
        private Integer age;

        @ExcelField(value = "时间", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private LocalDateTime time;

        @ExcelField(value = "excel时间", type = ExcelField.CellType.DATE, datePattern = "h:mm:ss")
        private double excelTime;

        @ExcelField(value = "字符串日期", type = ExcelField.CellType.DATE, datePattern = "yyyy-MM-dd")
        private String dateStr;

    }

    public static class CustomFontStyle implements FontStyle {

        @Override
        public boolean getItalic() {
            return true;
        }

        @Override
        public boolean getBold() {
            return true;
        }

        @Override
        public short getFontHeight() {
            return 256;
        }

        @Override
        public boolean getStrikeout() {
            return true;
        }
    }

}