package ziv.excel.news.invoker.impl;

import ziv.excel.news.invoker.PoiFiledInvoker;

import java.lang.reflect.Field;

/**
 * 默认的根据名称匹配的字段处理器
 * @author liuliuliu
 * @since 2021/10/28
 */
public abstract class PoiBaseFiledByNameInvoker implements PoiFiledInvoker {

    /**
     * 获取字段名
     * @return 会根据字段名进行匹配
     */
    public abstract String getFiledName();


    @Override
    public final boolean match(Object filed) {
        String filedName = getFiledName();
        if (filed == null || filedName == null)
            return false;
        return filed instanceof Field
                ? filedName.equalsIgnoreCase(((Field) filed).getName())
                : filedName.equalsIgnoreCase(filed.toString());
    }
}
