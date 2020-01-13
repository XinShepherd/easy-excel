package io.github.xinshepherd.excel.core;

/**
 * @author Fuxin
 * @since 2019/11/23 10:43
 */
public class ExcelException extends RuntimeException {

    public ExcelException() {
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }
}
