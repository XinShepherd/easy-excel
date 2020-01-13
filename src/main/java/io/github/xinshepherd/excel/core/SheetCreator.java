package io.github.xinshepherd.excel.core;

import io.github.xinshepherd.excel.annotation.Excel;
import io.github.xinshepherd.excel.annotation.ExcelBigHead;
import io.github.xinshepherd.excel.annotation.ExcelField;
import io.github.xinshepherd.excel.core.util.DateTimeUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.temporal.TemporalAccessor;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author Fuxin
 * @since 2020/1/13 15:54
 */
public class SheetCreator<T> {

    protected static final ExcelContext context = new ExcelContext();

    private final ExcelSheetMetadata<T> metadata;

    public SheetCreator(ExcelSheetMetadata<T> metadata) {
        this.metadata =  Objects.requireNonNull(metadata, "metadata不可为空");
    }

    public Sheet createSheet() {
        // 创建Sheet页
        Sheet sheet = metadata.getWorkbook().createSheet(metadata.getSheetName());
        AtomicInteger rowNumber = new AtomicInteger(0);
        handleBigHead(sheet, rowNumber);
        handleHeader(sheet, rowNumber);
        handleRows(sheet, rowNumber);
        return sheet;
    }

    protected  void handleBigHead(Sheet sheet, AtomicInteger rowNumber) {
        ExcelBigHead bigHead = metadata.getExcelBigHead();
        if (Objects.nonNull(bigHead)) {
            // 处理样式
            Class<? extends FontStyle> fontStyleClazz = bigHead.fontStyle();
            Workbook workbook = metadata.getWorkbook();
            Font font = createFont(workbook, context.getFontStyle(fontStyleClazz));
            CellStyle cellStyle = getMediumCellStyle(workbook);
            cellStyle.setFont(font);
            cellStyle.setAlignment(bigHead.horizontalAlignment());
            cellStyle.setVerticalAlignment(bigHead.verticalAlignment());

            Row row = sheet.createRow(bigHead.fromRow());
            Cell cell = row.createCell(bigHead.fromColumn());
            cell.setCellStyle(cellStyle);
            cell.setCellValue(bigHead.value());
            int toColumn = bigHead.autoMergeColumn()
                    ? Math.max(metadata.getExcelFields().size() - 1, 1)
                    : bigHead.toColumn();
            sheet.addMergedRegion(new CellRangeAddress(bigHead.fromRow(), bigHead.toRow(), bigHead.fromColumn(), toColumn));
            rowNumber.incrementAndGet();
        }
    }

    protected void handleHeader(Sheet sheet, AtomicInteger rowNumber) {
        Row header = sheet.createRow(rowNumber.get());
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
            CellStyle cellStyle = getMediumCellStyle(workbook);
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

    protected CellStyle getMediumCellStyle(Workbook workbook) {
        // 定义Cell格式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        return cellStyle;
    }

    protected void handleRows(Sheet sheet, AtomicInteger rowNumber) {
        try {
            List<Field> excelFields = metadata.getExcelFields();
            Excel metaExcel = metadata.getMetaExcel();
            for (int i = 0; i < metadata.getData().size(); i++) {
                T item = metadata.getData().get(i);
                Row row = sheet.createRow(rowNumber.incrementAndGet());
                row.setHeight(metaExcel.rowHigh());
                for (int j = 0; j < excelFields.size(); j++) {
                    Cell cell = row.createCell(j);
                    Field field = excelFields.get(j);
                    Object o = field.get(item);
                    handleCells(cell, field, o);
                }
            }
        } catch (Exception e) {
            throw new ExcelException(e);
        }
    }

    private static final Set<ExcelField.CellType> DATE_CELL_TYPES = EnumSet.of(ExcelField.CellType.DATE, ExcelField.CellType.TIME);
    protected void handleCells(Cell cell,
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
                    setDateValue(cell, fieldValue);
                    break;
                case TIME:
                    setTimeValue(cell, fieldValue);
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

    private void setTimeValue(Cell cell, Object fieldValue) {
        if (fieldValue instanceof Date) {
            cell.setCellValue((Date) fieldValue);
        } else if (fieldValue instanceof Number) {
            cell.setCellValue(Double.valueOf(fieldValue.toString()));
        } else if (fieldValue instanceof TemporalAccessor) {
            cell.setCellValue(DateTimeUtil.convertTime((TemporalAccessor) fieldValue));
        } else if (fieldValue instanceof String) {
            cell.setCellValue(DateTimeUtil.convertTime((String) fieldValue));
        }
    }

    private void setDateValue(Cell cell, Object fieldValue) {
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
