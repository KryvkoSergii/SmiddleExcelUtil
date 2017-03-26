package ua.com.smiddle.excelutil;

import java.util.Date;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Created by ksa on 26.03.17.
 */
public class ExcelExporterFactory {

    private static final Map<Class, String> defaultTypeFormat = new ConcurrentHashMap<>();

    static {
        defaultTypeFormat.put(Integer.class, "#");
        defaultTypeFormat.put(Long.class, "#");
        defaultTypeFormat.put(Short.class, "#");
        defaultTypeFormat.put(Byte.class, "#");
        defaultTypeFormat.put(Double.class, "#.##");
        defaultTypeFormat.put(Float.class, "#.##");
        defaultTypeFormat.put(Date.class, "dd/mm/yyyy");
    }

    public static ExcelExporter buildNewInstance() {
        return new ExcelExporter(defaultTypeFormat, Configurer.buildNewConfigurer());
    }

    public static ExcelExporter buildNewInstance(Configurer configurer) {
        return new ExcelExporter(defaultTypeFormat, configurer);
    }
}
