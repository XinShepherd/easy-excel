package cn.shepherd.excel.core.base;

import cn.shepherd.excel.annotation.Excel;
import cn.shepherd.excel.annotation.ExcelField;
import cn.shepherd.excel.core.*;
import cn.shepherd.excel.core.util.DateTimeUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.temporal.TemporalAccessor;
import java.util.*;

/**
 * @author Fuxin
 * @since 2019/11/23 10:44
 */
public abstract class ExporterBase {

    protected final ExcelContext context = new ExcelContext();

    @SuppressWarnings("unchecked")
    public Workbook generateWorkBook(Map<Class, List> dataMap) {
        Workbook workbook = new HSSFWorkbook();
        dataMap.forEach((key, value) -> appendSheet(key, value, workbook));
        return workbook;
    }


    public <T> ExcelMetadata<T> appendSheet(Class<T> clazz, List<T> data, Workbook workbook) {
        ExcelMetadata<T> metadata = new ExcelMetadata<>(clazz, data, workbook);
        createSheet(metadata);
        return metadata;
    }

    public <T> ExcelMetadata<T> generateMetadata(Class<T> clazz, List<T> data) {
        ExcelMetadata<T> metadata = new ExcelMetadata<>(clazz, data);
        createSheet(metadata);
        return metadata;
    }

    private <T> void createSheet(ExcelMetadata<T> metadata) {
        // 创建Sheet页
        Sheet sheet = metadata.getWorkbook().createSheet(metadata.getMetaExcel().value());
        handleHeader(metadata, sheet);
        handleRows(metadata, sheet);
    }

    private <T> void handleRows(ExcelMetadata<T> metadata, Sheet sheet) {
        try {
            List<Field> excelFields = metadata.getExcelFields();
            Excel metaExcel = metadata.getMetaExcel();
            for (int i = 0; i < metadata.getData().size(); i++) {
                T item = metadata.getData().get(i);
                Row row = sheet.createRow(i + 1);
                row.setHeight(metaExcel.rowHigh());
                for (int j = 0; j < excelFields.size(); j++) {
                    Cell cell = row.createCell(j);
                    Field field = excelFields.get(j);
                    Object o = field.get(item);
                    handleCells(metadata, cell, field, o);
                }
            }
        } catch (Exception e) {
            throw new ExcelException(e);
        }
    }

    private static final Set<ExcelField.CellType> DATE_CELL_TYPES = EnumSet.of(ExcelField.CellType.DATE, ExcelField.CellType.TIME);
    private <T> void handleCells(ExcelMetadata<T> metadata,
                                 Cell cell,
                                 Field field,
                                 Object fieldValue) {
        ExcelField excelField = field.getAnnotation(ExcelField.class);
        if (Objects.nonNull(fieldValue)) {
            // 定义Cell格式
            CellStyle cellStyle = metadata.getCellStyleMap().computeIfAbsent(field.getName(), key -> {
                CellStyle style = metadata.getWorkbook().createCellStyle();
                style.setAlignment(excelField.horizontalAlignment());
                style.setVerticalAlignment(excelField.verticalAlignment());
                if (DATE_CELL_TYPES.contains(excelField.type())) {
                    style.setDataFormat(metadata.getCreationHelper().createDataFormat().getFormat(excelField.datePattern()));
                }
                return style;
            });
            switch (excelField.type()) {
                case DATE:
                    if (fieldValue instanceof Date) {
                        cell.setCellValue((Date) fieldValue);
                    } else if (fieldValue instanceof Number) {
                        cell.setCellValue(Double.valueOf(fieldValue.toString()));
                    } else if (fieldValue instanceof LocalDateTime) {
                        cell.setCellValue((LocalDateTime) fieldValue);
                    } else if (fieldValue instanceof LocalDate) {
                        cell.setCellValue((LocalDate) fieldValue);
                    } else if (fieldValue instanceof String) {
                        Double excelTime = DateTimeUtil.parseDateTime((String) fieldValue);
                        if(Objects.nonNull(excelTime))
                            cell.setCellValue(excelTime);
                    }
                    break;
                case TIME:
                    if (fieldValue instanceof Date) {
                        cell.setCellValue((Date) fieldValue);
                    } else if (fieldValue instanceof Number) {
                        cell.setCellValue(Double.valueOf(fieldValue.toString()));
                    } else if (fieldValue instanceof TemporalAccessor) {
                        cell.setCellValue(DateTimeUtil.convertTime((TemporalAccessor) fieldValue));
                    } else if (fieldValue instanceof String) {
                        cell.setCellValue(DateTimeUtil.convertTime((String) fieldValue));
                    }
                    break;
                case NUMERIC:
                    cell.setCellValue(Double.valueOf(fieldValue.toString()));
                    break;
                default: // 默认字符串格式
                    cell.setCellValue(fieldValue.toString());
            }
            cell.setCellStyle(cellStyle);
        } else {
            cell.setBlank();
        }
    }

    private <T> void handleHeader(ExcelMetadata<T> metadata, Sheet sheet) {
        Row header = sheet.createRow(0);
        List<Field> excelFields = metadata.getExcelFields();
        Excel metaExcel = metadata.getMetaExcel();
        Workbook workbook = metadata.getWorkbook();
        // 处理字体
        Class<? extends FontStyle> fontStyleClazz = metaExcel.fontStyle();
        Font font = createFont(workbook, context.getFontStyle(fontStyleClazz));
        // 处理表头高度
        header.setHeight(metaExcel.herderHigh());
        // 处理是否固定表头
        if (metaExcel.freezePane()) {
            sheet.createFreezePane(1, 1);
        }
        for (int i = 0; i < excelFields.size(); i++) {
            Field field = excelFields.get(i);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            Cell cell = header.createCell(i);
            cell.setCellValue(excelField.value());
            // 定义Cell格式
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            cellStyle.setBorderRight(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            Class<? extends FontStyle> cellFontStyle = excelField.fontStyle();
            if (DefaultFontStyle.class.equals(cellFontStyle)) {
                cellStyle.setFont(font);
            } else {
                Font cellFont = createFont(workbook, context.getFontStyle(cellFontStyle));
                cellStyle.setFont(cellFont);
            }
            if (metaExcel.headerColor() != -1) {
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(metaExcel.headerColor());
            }
            if (excelField.headerColor() != -1) {
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(excelField.headerColor());
            }
            cellStyle.setAlignment(excelField.horizontalAlignment());
            cellStyle.setVerticalAlignment(excelField.verticalAlignment());
            cell.setCellStyle(cellStyle);
            sheet.setColumnWidth(i, excelField.width() * 128);
        }
    }

    private Font createFont(Workbook workbook, FontStyle fontStyle) {
        Font font = workbook.createFont();
        font.setBold(fontStyle.getBold());
        font.setItalic(fontStyle.getItalic());
        font.setUnderline(fontStyle.getUnderline());
        font.setTypeOffset(fontStyle.getTypeOffset());
        font.setColor(fontStyle.getColor());
        font.setStrikeout(fontStyle.getStrikeout());
        font.setCharSet(fontStyle.getCharSet());
        if (Objects.nonNull(fontStyle.getFontName())) {
            font.setFontName(fontStyle.getFontName());
        }
        if (fontStyle.getFontHeight() != -1) {
            font.setFontHeight(fontStyle.getFontHeight());
        }
        if (fontStyle.getFontHeightInPoints() != -1) {
            font.setFontHeight(fontStyle.getFontHeightInPoints());
        }
        return font;
    }
}
