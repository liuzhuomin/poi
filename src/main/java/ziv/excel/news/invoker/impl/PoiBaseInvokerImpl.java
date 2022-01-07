package ziv.excel.news.invoker.impl;

import ziv.excel.news.invoker.PoiCommentInvoker;
import ziv.excel.news.invoker.PoiFiledInvoker;
import ziv.excel.news.invoker.PoiInvoker;
import ziv.excel.news.invoker.PoiStyleInvoker;

import java.util.List;


/**
 * 基础invoker实现
 *
 * @param <T>
 * @author liuliuliu
 * @since 2021/1/20
 */
public abstract class PoiBaseInvokerImpl<T> implements PoiInvoker<T> {
    protected PoiCommentInvoker poiCommentInvoker;
    protected PoiStyleInvoker poiStyleInvoker;
    protected List<PoiFiledInvoker> poiFiledInvokerList;
    protected String[] ignore;

    public PoiCommentInvoker getPoiCommentInvoker() {
        return poiCommentInvoker;
    }

    @Override
    public PoiStyleInvoker getPoiStyleInvoker() {
        return poiStyleInvoker;
    }

    public List<PoiFiledInvoker> getPoiFiledInvokerList() {
        return poiFiledInvokerList;
    }

    public String[] getIgnore() {
        return ignore;
    }

    public PoiBaseInvokerImpl<T> setPoiCommentInvoker(PoiCommentInvoker poiCommentInvoker) {
        this.poiCommentInvoker = poiCommentInvoker;
        return this;
    }

    public PoiBaseInvokerImpl<T> setPoiStyleInvoker(PoiStyleInvoker poiStyleInvoker) {
        this.poiStyleInvoker = poiStyleInvoker;
        return this;
    }

    public PoiBaseInvokerImpl<T> setPoiFiledInvokerList(List<PoiFiledInvoker> poiFiledInvokerList) {
        this.poiFiledInvokerList = poiFiledInvokerList;
        return this;
    }

    public PoiBaseInvokerImpl<T> setIgnore(String[] ignore) {
        this.ignore = ignore;
        return this;
    }
}
