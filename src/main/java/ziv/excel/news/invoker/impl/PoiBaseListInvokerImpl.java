package ziv.excel.news.invoker.impl;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ArrayUtil;
import ziv.excel.news.invoker.PoiFiledInvoker;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.lang.reflect.ParameterizedType;
import java.util.*;
import java.util.function.Consumer;

/**
 * list转book实现
 * book转list实现
 * 处理list类型的
 *
 * @param <S>
 * @author liuliuliu
 * @since 2021/1/20
 */
public class PoiBaseListInvokerImpl<S> extends PoiBaseInvokerImpl<List<S>> {

    private Class<S> clazz;

    @Override
    public TreeMap<String, String> getHeaders(List<S> source) {
        Assert.notNull(poiCommentInvoker, "poiCommentInvoker must not be null!");
        TreeMap<String, String> result = new TreeMap<>();
        Class<?> aClass = getClazz(source);
        //反射映射，并且通过定义工具取注释说明
        filter(aClass, field -> result.put(field.getName(), poiCommentInvoker.getComment(field)));
        return result;
    }

    @Override
    public List<Map<String, Object>> getRows(List<S> source) {
        List<Map<String, Object>> result = new ArrayList<>();
        if (CollectionUtil.isNotEmpty(source)) {
            Class<?> aClass = getClazz(source);
            for (S s : source) {
                Map<String, Object> m = new HashMap<>();
                filter(aClass, field -> {
                    Object filedVar = getFiledVar(s, field);
                    if (CollectionUtil.isNotEmpty(poiFiledInvokerList)) {
                        PoiFiledInvoker poiFiledInvoker = poiFiledInvokerList
                                .stream()
                                .filter(inner -> inner.match(field))
                                .findAny()
                                .orElse(null);
                        if (poiFiledInvoker != null) {
                            filedVar = poiFiledInvoker.getFiledVar(filedVar);
                        }
                    }
                    m.put(field.getName(), filedVar);
                });
                result.add(m);
            }
        }
        return result;
    }

    private Object getFiledVar(S s, Field field) {
        Object filedVar = null;
        try {
            filedVar = field.get(s);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return filedVar;
    }

    private void filter(Class<?> aClass, Consumer<Field> consumer) {
        Field[] declaredFields = aClass.getDeclaredFields();
        for (Field declaredField : declaredFields) {
            declaredField.setAccessible(true);
            String name = declaredField.getName();
            if (ArrayUtil.isEmpty(ignore) || !ArrayUtil.contains(ignore, name)) {
                int modifiers = declaredField.getModifiers();
                if (!Modifier.isFinal(modifiers)) {
                    consumer.accept(declaredField);
                }
            }
        }
    }

    private Class<?> getClazz(List<S> source) {
        if (this.clazz != null) {
            return clazz;
        }
        //解析class
        Class<?> aClass;
        if (CollectionUtil.isNotEmpty(source)) {
            aClass = source.get(0).getClass();
        } else {
            ParameterizedType parameterizedType = (ParameterizedType) this.getClass().getGenericSuperclass();
            aClass = (Class<?>) parameterizedType.getActualTypeArguments()[0];
        }
        return aClass;
    }

    public Class<S> getClazz() {
        return clazz;
    }

    public void setClazz(Class<S> clazz) {
        this.clazz = clazz;
    }
}
