package ziv.excel.news.invoker;


import ziv.excel.news.invoker.impl.*;

import java.util.List;
import java.util.Map;

/**
 * invoker生成器
 *
 * @author liuliuliu
 * @since 2021/1/20
 */
public class PoiInvokerFactory {

    /**
     * 生成的是默认转list的
     *
     * @param poiFiledInvokers 特殊字段
     * @param ignoreFields     忽略字段
     * @param <T>              任意类型
     * @return poi处理器
     */
    public static <T> PoiInvoker<List<T>> listInvoker(List<PoiFiledInvoker> poiFiledInvokers,
                                                      String... ignoreFields) {
        PoiBaseListInvokerImpl<T> poiInvoker = new PoiBaseListInvokerImpl<>();
        poiInvoker.setIgnore(ignoreFields);
        poiInvoker.setPoiCommentInvoker(new PoiCommentBySwaggerInvoker());
        poiInvoker.setPoiStyleInvoker(new PoiBaseStyleInvokerImpl());
        poiInvoker.setPoiFiledInvokerList(poiFiledInvokers);
        return poiInvoker;
    }


    /**
     * 生成的是默认转map的
     *
     * @param poiFiledInvokers 特殊字段
     * @param ceTable          key英文，value中文（用来对照map中的英文）
     * @param ignoreFields     忽略字段
     * @param <T>              任意类型
     * @return poi处理器
     */
    public static <T> PoiInvoker<List<T>> mapInvoker(List<PoiFiledInvoker> poiFiledInvokers,
                                                     Map<String, String> ceTable,
                                                     String... ignoreFields) {
        PoiBaseMapInvokerImpl<T> poiInvoker = new PoiBaseMapInvokerImpl<>();
        poiInvoker.setIgnore(ignoreFields);
        poiInvoker.setPoiFiledInvokerList(poiFiledInvokers);
        poiInvoker.setPoiCommentInvoker(new PoiCommentByMapInvoker(ceTable));
        poiInvoker.setPoiStyleInvoker(new PoiBaseStyleInvokerImpl());
        return poiInvoker;
    }



}
