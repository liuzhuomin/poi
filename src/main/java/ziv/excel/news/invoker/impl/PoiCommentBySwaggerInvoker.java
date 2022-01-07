package ziv.excel.news.invoker.impl;


import ziv.excel.news.invoker.PoiCommentInvoker;
import io.swagger.annotations.ApiModelProperty;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;

/**
 * 根据swagger的注解{@link ApiModelProperty}获取字段上的值
 * <li>
 * 如果对象字段上标注了{@link ApiModelProperty}注解，则取注解的{@link ApiModelProperty#value()}函数
 * 如果{@link ApiModelProperty#value()}的值为空，取{@link ApiModelProperty#name()}函数返回的值
 * </li>
 * <li>
 * 如果对象字段上未曾标注{@link ApiModelProperty}，则直接取字段名作为值
 * </li>
 *
 * @author liuliuliu
 * @since 2021/1/20
 */
public class PoiCommentBySwaggerInvoker implements PoiCommentInvoker {

	/**
	 * 获取字段对应的描述信息，也就是字段的释义，如果无注解或者注解无值，取字段名为释义
	 * @param sourceField 字段
	 * @return 注释值
	 */
    @Override
    public String getComment(Object sourceField) {
        String key = null;
        if (sourceField instanceof Field) {
            Field field = (Field) sourceField;
            boolean annotationPresent = field.isAnnotationPresent(ApiModelProperty.class);
            if (annotationPresent) {
                ApiModelProperty annotation = field.getAnnotation(ApiModelProperty.class);
                key = annotation.value();
                if (StringUtils.isBlank(key)) {
                    key = field.getName();
                }
            } else {
                key = field.getName();
            }
        }
        return key;
    }
}
