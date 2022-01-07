package ziv.excel.news.invoker.impl;

import ziv.excel.news.invoker.PoiFiledInvoker;

import java.lang.reflect.Field;

/**
 * 默认的根据名称匹配的字段处理器
 *
 * @author liuliuliu
 * @since 2021/10/28
 */
public abstract class PoiBaseFiledByClassInvoker implements PoiFiledInvoker {

    /**
     * 获取字段名
     *
     * @return 会根据字段类型进行匹配
     */
    public abstract Class getType();


    @Override
    public final boolean match(Object filed) {
        Class filedName = getType();
        if (filed == null || filedName == null)
            return false;
        return filed instanceof Field
                ? filedName == (((Field) filed).getType())
                : filedName == filed.getClass();
    }
}
