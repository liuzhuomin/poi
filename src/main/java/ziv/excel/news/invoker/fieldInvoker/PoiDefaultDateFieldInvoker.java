package ziv.excel.news.invoker.fieldInvoker;

import ziv.excel.news.invoker.impl.PoiBaseFiledByClassInvoker;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 默认的Date字段处理器，时间格式为"yyyy-MM-dd HH:mm:ss"
 *
 * @author liuliuliu
 * @since 2021/10/28
 */
public class PoiDefaultDateFieldInvoker extends PoiBaseFiledByClassInvoker {

    @Override
    public Class getType() {
        return Date.class;
    }

    @Override
    public Object getFiledVar(Object obj) {
        return obj == null ? null : new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format((Date) obj);
    }
}
