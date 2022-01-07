package ziv.excel.news.invoker.impl;

import cn.hutool.core.lang.Assert;
import ziv.excel.news.invoker.PoiCommentInvoker;
import org.apache.commons.lang3.StringUtils;

import java.util.Map;
import java.util.logging.Logger;

/**
 * 根据config的map获取字段对应的值，如果为空则返回null
 *
 * @author liuliuliu
 * @since 2021/1/20
 */
public class PoiCommentByMapInvoker implements PoiCommentInvoker {

    Map<String, String> config;
    public PoiCommentByMapInvoker(Map<String, String> config) {
        this.config = config;
    }

    /**
     * 根据config的map获取字段对应的值，如果为空则返回null
     * 以<b>sourceField</b>的{@link Object#toString()}函数作为key，取config中对应的value
     * @param sourceField 字段
     * @return 字段对应的注释
     */
    @Override
    public String getComment(Object sourceField) {
        Assert.notNull(sourceField, "sourceField must not be null!");
        String s = config.get(sourceField.toString());
        if (StringUtils.isBlank(s)) {
            Logger.getGlobal().info("未取到值!");
            return null;
        }
        return s;
    }
}
