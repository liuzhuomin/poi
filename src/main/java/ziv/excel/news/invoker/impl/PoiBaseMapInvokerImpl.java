package ziv.excel.news.invoker.impl;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ArrayUtil;
import org.apache.commons.lang3.StringUtils;
import ziv.excel.news.PoiConfig;

import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

/**
 * map转book实现
 * book转map实现
 * 处理map类型的
 *
 * @param <S>
 * @author liuliuliu
 * @since 2021/1/20
 */
public class PoiBaseMapInvokerImpl<S> extends PoiBaseInvokerImpl<List<S>> {


    @Override
    public Map<String, String> getHeaders(List<S> source) {

        Assert.notNull(poiCommentInvoker, "poiCommentInvoker must not be null!");
        TreeMap<String, String> result = new TreeMap<>();

        for (S s : source) {
            Map s1 = (Map) s;
            Set set = s1.keySet();

            for (Object o : set) {
                if (ArrayUtil.isEmpty(ignore) || !ArrayUtil.contains(ignore, o.toString())) {
                    String comment = poiCommentInvoker.getComment(o);
                    //强制性忽略，不做兼容
                    if (PoiConfig.EXPORT_AUTO_FIELD || StringUtils.isNotBlank(comment))
                        result.put(o.toString(), comment);
                }
            }
            break;
        }
        return result;
    }

    @Override
    public List<Map<String, Object>> getRows(List<S> source) {
        List<Map<String, Object>> result = new ArrayList<>();
        if (CollectionUtil.isNotEmpty(source)) {
            for (S s : source) {
                Map prototype = (Map) s;
                Map<String, Object> singleResult = new HashMap<>();
                boolean notEmpty = CollectionUtil.isNotEmpty(poiFiledInvokerList);

                for (Object key : prototype.keySet()) {
                    String keyObj = key.toString();
                    if (ArrayUtil.isEmpty(ignore) || !ArrayUtil.contains(ignore, keyObj)) {
                        AtomicReference<Object> valueObj = new AtomicReference<>(prototype.get(key));
                        if (notEmpty) {
                            poiFiledInvokerList
                                    .stream()
                                    .filter(inner -> inner.match(keyObj))
                                    .findAny().ifPresent(poiFiledInvoker -> {
                                valueObj.set(poiFiledInvoker.getFiledVar(valueObj.get()));
                            });
                        }
                        singleResult.put(keyObj, valueObj.get());
                    }
                }
                if (!singleResult.isEmpty()) {
                    result.add(singleResult);
                }
            }
        }
        return result;
    }

}
