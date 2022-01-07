package ziv.excel.news;

/**
 * poi配置
 */
public class PoiConfig {
    /**
     * 自动生成字段，不匹配配置项和数据项是否匹配
     */
    public static boolean EXPORT_AUTO_FIELD = true;

    /**
     * 是否校验整个导入的excel，如果一行不对直接抛出异常
     */
    public static boolean IMPORT_VALID_WHOLE_EXCEL = true;
}
