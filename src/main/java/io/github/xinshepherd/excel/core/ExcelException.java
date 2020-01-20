package io.github.xinshepherd.excel.core;

/**
 * @author Fuxin
 * @since 1.1.0
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
