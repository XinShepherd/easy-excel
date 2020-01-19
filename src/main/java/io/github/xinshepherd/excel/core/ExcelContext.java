package io.github.xinshepherd.excel.core;

import java.lang.reflect.Constructor;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author Fuxin
 * @since 2019/11/19 9:39
 */

public class ExcelContext {

    /**
     * 根据全类名获取字体类型bean
     */
    private final Map<String, FontStyle> fontStyleMap;

    /** Cache of singleton objects created by FactoryBeans: FactoryBean name to object. */
    private final Map<String, Object> factoryBeanObjectCache = new ConcurrentHashMap<>(16);

    public ExcelContext() {
        fontStyleMap = new ConcurrentHashMap<>();
    }

    public FontStyle getFontStyle(Class<? extends FontStyle> clazz) {
        return fontStyleMap.computeIfAbsent(clazz.getName(), key -> {
            try {
                Constructor<? extends FontStyle> constructor = clazz.getConstructor();
                constructor.setAccessible(true);
                return constructor.newInstance();
            } catch (Exception e) {
                throw new ExcelException(e);
            }
        });
    }

    @SuppressWarnings("unchecked")
    public <T> T getBean(Class<T> clazz) {
        return (T) factoryBeanObjectCache.computeIfAbsent(clazz.getName(), key -> {
            try {
                Constructor<T> constructor = clazz.getConstructor();
                constructor.setAccessible(true);
                return constructor.newInstance();
            } catch (Exception e) {
                throw new ExcelException(e);
            }
        });
    }
}
