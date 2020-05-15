package io.github.xinshepherd.excel.core.base;

import io.github.xinshepherd.excel.annotation.Excel;
import io.github.xinshepherd.excel.annotation.ExcelField;
import io.github.xinshepherd.excel.core.ExcelException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;
import java.util.function.Function;

/**
 * 将excel表格导入解析成对应的java类
 *
 * @author donglin
 * @date 2020/05/15
 */
public class ImporterBase {

    private static final String FILE_NAME_SUFFIX_XLS = ".xls";
    private static final String FILE_NAME_SUFFIX_XLSX = ".xlsx";

    private static final String CONTEXT_TYPE_XLS = "application/vnd.ms-excel";
    private static final String CONTEXT_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /**
     * 根据标题来匹配. 即通过ExcelField注解的value值来寻找excel表格里的列标题所在的列，相同的话就将那一列映射到对应的java类字段
     *
     * @see ExcelField value
     */
    private static final String MATCH_TYPE_TITLE = "TITLE";

    /**
     * 根据位置来匹配.  即通过ExcelField注解的position值来寻找excel表格里第几列(0开头)，相同的话就将那一列映射到对应的java类字段
     *
     * @see ExcelField position
     */
    private static final String MATCH_TYPE_POSITION = "POSITION";

    protected static final Function<Cell, Object> CONVERT_NUMBER = Cell::getNumericCellValue;
    protected static final Function<Cell, Object> CONVERT_INTEGER = cell -> (int) cell.getNumericCellValue();
    protected static final Function<Cell, Object> CONVERT_SHORT = cell -> (int) cell.getNumericCellValue();
    protected static final Function<Cell, Object> CONVERT_LONG = cell -> (long) cell.getNumericCellValue();
    protected static final Function<Cell, Object> CONVERT_DOUBLE = Cell::getNumericCellValue;
    protected static final Function<Cell, Object> CONVERT_FLOAT = cell -> (float) cell.getNumericCellValue();
    protected static final Function<Cell, Object> CONVERT_STRING = Cell::getStringCellValue;
    protected static final Function<Cell, Object> CONVERT_DATE = Cell::getDateCellValue;

    /**
     * 映射类型，默认是按标题的
     */
    private String matchType = MATCH_TYPE_TITLE;

    /**
     * 文件格式. 优先用这个来判断是否excel文件，且是xls还是xlxs
     */
    private String contextType;

    /**
     * 文件名. 没有contextType的话则会通过文件名后缀来判断文件类型
     */
    private String filename = "";

    /**
     * 列标题所在的行号，0起始，默认0
     */
    private int titleRowIndex = 0;

    /**
     * 倒数第 ignoreLastIndexes 行开始都忽略，有些excel表格最后几行会是一些统计信息，无法与前面列标题匹配，可以忽略掉
     */
    private int ignoreLastIndexes = 0;

    /**
     * 输入流
     */
    private final InputStream inputStream;


    /**
     * key: excel表格的第几列 (从0开始)
     * value: 对应的java类字段
     */
    private Map<Integer, Field> columnFieldMap = new HashMap<>();

    /**
     * 数据格式转换器MAP
     */
    private Map<Field, Function<Cell, Object>> dataCovertMap = new HashMap<>();


    public static ImporterBase newInstance(InputStream inputStream) {
        if (inputStream == null) {
            throw new NullPointerException("inputStream not null");
        }
        return new ImporterBase(inputStream);
    }

    private ImporterBase(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public ImporterBase matchType(String matchType) {
        this.matchType = matchType;
        return this;
    }

    public ImporterBase contextType(String contextType) {
        this.contextType = contextType;
        return this;
    }

    public ImporterBase filename(String filename) {
        this.filename = filename;
        return this;
    }

    public ImporterBase titleRowIndex(int titleRowIndex) {
        this.titleRowIndex = titleRowIndex;
        return this;
    }

    public ImporterBase ignoreLastIndexes(int ignoreLastIndexes) {
        this.ignoreLastIndexes = ignoreLastIndexes;
        return this;
    }

    public <T> List<T> resolve(Class<T> cls) throws Exception {
        Annotation annotation = cls.getAnnotation(Excel.class);
        if (annotation == null) {
            throw new ExcelException(cls.getName() + " not annotation " + Excel.class.getName());
        }
        Workbook workbook;
        if (contextType != null && !"".equals(contextType)) {
            workbook = newWorkbookByContextType(inputStream, contextType);
        } else if (filename != null && !"".equals(filename)) {
            workbook = newWorkbookByFilename(inputStream, filename);
        } else {
            workbook = newWorkbookByContextType(inputStream, CONTEXT_TYPE_XLSX);
        }
        if (MATCH_TYPE_POSITION.equals(matchType)) {
            return resolveByPosition(workbook, cls);
        } else {
            return resolveByTitle(workbook, cls);
        }
    }

    private <T> List<T> resolveByTitle(Workbook workbook, Class<T> cls) throws Exception {
        try {
            List<T> list = new ArrayList<>();
            // 暂时只处理第一个sheet
            Sheet sheet = workbook.getSheetAt(0);
            // 获取标题行
            Row titleRow = sheet.getRow(this.titleRowIndex);

            initColumnFieldMap(titleRow, cls);

            // 从标题行的下一行开始解析，并忽略掉最后几行需要忽略的
            int start = titleRowIndex + 1;
            int end = sheet.getPhysicalNumberOfRows() - ignoreLastIndexes;

            int cellNumber = titleRow.getPhysicalNumberOfCells();
            for (int i = start; i < end; i++) {
                titleRow = sheet.getRow(i);
                T t = cls.newInstance();
                for (int j = 0; j < cellNumber; j++) {
                    Field field = columnFieldMap.get(j);
                    if (field != null && titleRow.getCell(j) != null) {
                        writeValue(t, field, titleRow.getCell(j));
                    }
                }
                list.add(t);
            }
            return list;
        } finally {
            workbook.close();
        }
    }

    private <T> List<T> resolveByPosition(Workbook workbook, Class<T> cls) {
        throw new ExcelException("Not support MATCH_TYPE_POSITION for now");
    }

    private <T> void writeValue(T t, Field field, Cell cell) {
        if (CellType._NONE.equals(cell.getCellType())) {
            return;
        }
        try {
            field.set(t, dataCovertMap.get(field).apply(cell));
        } catch (Exception e) {
            throw new ExcelException(String.format("%s covert to %s error.", cell.getStringCellValue(), field.getName()));
        }
    }


    private <T> void initColumnFieldMap(Row row, Class<T> cls) {
        int size = row.getPhysicalNumberOfCells();
        columnFieldMap = new HashMap<>(size);
        Map<String, Integer> indexMap = new HashMap<>(size);
        for (int i = 0; i < size; i++) {
            indexMap.put(row.getCell(i).getStringCellValue(), i);
        }
        for (Field field : cls.getDeclaredFields()) {
            ExcelField excelField = field.getDeclaredAnnotation(ExcelField.class);
            if (excelField != null) {
                String value = excelField.value();
                if (indexMap.containsKey(value)) {
                    field.setAccessible(true);
                    columnFieldMap.put(indexMap.get(value), field);
                }
            }
        }

        dataCovertMap = new HashMap<>(16);
        for (Field field : columnFieldMap.values()) {
            dataCovertMap.put(field, dataCovert(field));
        }
    }


    protected Function<Cell, Object> dataCovert(Field field) {
        Class<?> cls = field.getType();
        if (cls.equals(String.class)) {
            return CONVERT_STRING;
        }
        if (cls.equals(int.class) || cls.equals(Integer.class)) {
            return CONVERT_INTEGER;
        }
        if (cls.equals(long.class) || cls.equals(Long.class)) {
            return CONVERT_LONG;
        }
        if (cls.equals(short.class) || cls.equals(Short.class)) {
            return CONVERT_SHORT;
        }
        if (cls.equals(double.class) || cls.equals(Double.class)) {
            return CONVERT_DOUBLE;
        }
        if (cls.equals(float.class) || cls.equals(Float.class)) {
            return CONVERT_FLOAT;
        }
        if (cls.equals(Date.class)) {
            return CONVERT_DATE;
        }
        throw new ExcelException(String.format("Not support %s for now.", cls.getName()));
    }


    private static Workbook newWorkbookByFilename(InputStream in, String filename) throws IOException {
        String suffix = filename.contains(".") ? filename.substring(filename.lastIndexOf('.')) : "";
        if (FILE_NAME_SUFFIX_XLS.equalsIgnoreCase(suffix)) {
            return new HSSFWorkbook(in);
        } else if (FILE_NAME_SUFFIX_XLSX.equalsIgnoreCase(suffix)) {
            return new XSSFWorkbook(in);
        } else {
            throw new ExcelException("the file is not excel.");
        }
    }

    private static Workbook newWorkbookByContextType(InputStream in, String contextType) throws IOException {
        if (CONTEXT_TYPE_XLS.equalsIgnoreCase(contextType)) {
            return new HSSFWorkbook(in);
        } else if (CONTEXT_TYPE_XLSX.equalsIgnoreCase(contextType)) {
            return new XSSFWorkbook(in);
        } else {
            throw new ExcelException("the file is not excel.");
        }
    }

}
