package ziv.excel.news.imports;

import io.swagger.annotations.ApiModelProperty;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 根据swagger注解解析field
 *
 * @author liuliuliu
 * @since 2021/10/29
 */
public class ParseBySwagger {

    private static String[] IGNORE = {"serialVersionUID"};


    public static Map<String, Field> getFiledMap(Class<?> targetClass, String... ignores) {
        Map<String, Field> allFieldMap = new HashMap<>();
        Field[] declaredFields = getDeclareFields(targetClass, ignores);
        for (Field declaredField : declaredFields) {
            //处理header头
            boolean annotationPresent = declaredField.isAnnotationPresent(ApiModelProperty.class);
            if (annotationPresent) {
                ApiModelProperty annotation = declaredField.getAnnotation(ApiModelProperty.class);
                String name = annotation.name();
                if (StringUtils.isBlank(name)) {
                    allFieldMap.put(annotation.value(), declaredField);
                } else {
                    allFieldMap.put(name, declaredField);
                }
            } else {
                allFieldMap.put(declaredField.getName(), declaredField);
            }
        }
        return allFieldMap;
    }

    /**
     * 获取类的字段（忽略掉IGNORE的所有同名字段）
     *
     * @param clazz clazz
     * @return Field[]
     */
    static Field[] getDeclareFields(Class<?> clazz, String... ignores) {
        Field[] declaredFields = clazz.getDeclaredFields();
        List<Field> collect = Arrays.stream(declaredFields).collect(Collectors.toList());
        for (Field declaredField : declaredFields) {
            for (String s : IGNORE) {
                String name = declaredField.getName();
                if (name.equals(s)) {
                    collect.remove(declaredField);
                    break;
                }
            }
            for (String s : ignores) {
                String name = declaredField.getName();
                if (name.equals(s)) {
                    collect.remove(declaredField);
                    break;
                }
            }
        }
        return collect.toArray(new Field[0]);
    }

}
