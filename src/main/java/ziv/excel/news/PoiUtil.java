package ziv.excel.news;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.Workbook;
import ziv.excel.news.imports.PoiBookParser;
import ziv.excel.news.invoker.*;
import ziv.excel.news.invoker.impl.PoiBaseListInvokerImpl;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

/**
 * 导入、导出、生成模板
 * 附加功能: 字段处理
 * 附加功能: 字段扩展和减少
 * 附加功能: list导入导出,map导入导出,单个实体导入导出
 * 附加功能: 解析器，负责解析传入对象为具体下游对象
 * <p>
 * <p>
 * map转换器，list转换器
 * 字段处理器，模板生成器，header头生成器，行渲染器，列渲染器，数据校验器
 *
 * @author liuliuliu
 * @see PoiFiledInvoker
 * @see PoiStyleInvoker
 * @see PoiCommentInvoker
 * @see PoiInvoker
 * @since 2021/1/20
 */
public final class PoiUtil {

    /**
     * 根据class对象生成模板
     *
     * @param sheetName sheet页名称
     * @param clazz     类型
     * @param <T>       任意类型
     * @return {@link Workbook}对象
     */
    public static <T> Workbook generateTemplate(String sheetName, Class<T> clazz) {
        PoiInvoker<List<T>> listPoiInvoker = PoiInvokerFactory.listInvoker(null);
        ((PoiBaseListInvokerImpl) listPoiInvoker).setClazz(clazz);
        return new PoiIInvokerImpl<T>().generate(sheetName, Lists.newArrayList(), listPoiInvoker, null, null);
    }

    /**
     * 根据header生成模板
     *
     * @param sheetName sheet页名称
     * @param header    类型
     * @param <T>       任意类型
     * @return {@link Workbook}对象
     */
    public static <T> Workbook generateTemplate(String sheetName, Map<String, String> header) {
        return map2book(sheetName, Lists.newArrayList(), header);
    }


    /**
     * {@link List}转换成{@link Workbook}对象
     *
     * @param sheetName sheet页名称
     * @param dataList  数据集合
     * @param <T>       任意类型
     * @return {@link Workbook}对象
     */
    public static <T> Workbook list2book(String sheetName, List<T> dataList) {
        PoiInvoker<List<T>> listPoiInvoker = PoiInvokerFactory.listInvoker(null);
        return new PoiIInvokerImpl<T>().generate(sheetName, dataList, listPoiInvoker, null, null);
    }

    /**
     * {@link List}转换成{@link Workbook}对象
     *
     * @param sheetName    sheet页名称
     * @param dataList     数据集合
     * @param ignoreFields 忽略字段属性（匹配字段名）
     * @param <T>          任意类型
     * @return {@link Workbook}对象
     */
    public static <T> Workbook list2book(String sheetName, List<T> dataList, String... ignoreFields) {
        PoiInvoker<List<T>> listPoiInvoker = PoiInvokerFactory.listInvoker(null, ignoreFields);
        return new PoiIInvokerImpl<T>().generate(sheetName, dataList, listPoiInvoker, null, null);
    }

    /**
     * {@link List}转换成{@link Workbook}对象
     *
     * @param sheetName sheet页名称
     * @param dataList  数据集合
     * @param otherData 扩展数据，会在数据行的最下面隔开一行，排列生成
     * @param <T>       任意类型
     * @return {@link Workbook}对象
     */
    public static <S, T> Workbook list2bookWithOtherData(String sheetName, List<T> dataList, List<S> otherData) {
        PoiInvoker<List<T>> dataListInvoker = PoiInvokerFactory.listInvoker(null);
        PoiInvoker<List<S>> otherDataInvoker = PoiInvokerFactory.listInvoker(null);
        return new PoiIInvokerImpl<T>().generate(sheetName, dataList, dataListInvoker, otherData, otherDataInvoker);
    }

    /**
     * {@link List}转换成{@link Workbook}对象
     *
     * @param sheetName        sheet页名称
     * @param dataList         数据集合
     * @param poiFiledInvokers 字段处理器，需要处理字段的值时可使用，默认提供有
     *                         {@link ziv.excel.news.invoker.fieldInvoker.PoiDefaultDateFieldInvoker}
     *                         处理器
     * @param <T>              任意类型
     * @return {@link Workbook}对象
     */
    public static <T> Workbook list2bookWithInvokers(String sheetName, List<T> dataList, List<PoiFiledInvoker> poiFiledInvokers) {
        PoiInvoker<List<T>> dataListInvoker = PoiInvokerFactory.listInvoker(poiFiledInvokers);
        return new PoiIInvokerImpl<T>().generate(sheetName, dataList, dataListInvoker, null, null);
    }

    /**
     * {@link List}转换成{@link Workbook}对象
     *
     * @param sheetName    sheet页名称
     * @param dataList     数据集合
     * @param otherData    扩展数据，会在数据行的最下面隔开一行，排列生成
     * @param ignoreFields 忽略字段属性（匹配字段名）
     * @param <T>          任意类型
     * @return {@link Workbook}对象
     */
    public static <S, T> Workbook list2book(String sheetName, List<T> dataList, List<S> otherData, String... ignoreFields) {
        PoiInvoker<List<T>> dataListInvoker = PoiInvokerFactory.listInvoker(null, ignoreFields);
        PoiInvoker<List<S>> otherDataInvoker = PoiInvokerFactory.listInvoker(null, ignoreFields);
        return new PoiIInvokerImpl<T>().generate(sheetName, dataList, dataListInvoker, otherData, otherDataInvoker);
    }

    /**
     * {@link List}转换成{@link Workbook}对象
     *
     * @param sheetName        sheet页名称
     * @param dataList         数据集合
     * @param otherData        扩展数据，会在数据行的最下面隔开一行，排列生成
     * @param poiFiledInvokers 字段处理器，需要处理字段的值时可使用，默认提供有
     *                         {@link ziv.excel.news.invoker.fieldInvoker.PoiDefaultDateFieldInvoker}
     *                         处理器
     * @param ignoreFields     忽略字段属性（匹配字段名）
     * @param <T>              任意类型
     * @return {@link Workbook}对象
     */
    public static <S, T> Workbook list2book(String sheetName, List<T> dataList, List<S> otherData, List<PoiFiledInvoker> poiFiledInvokers, String... ignoreFields) {
        PoiInvoker<List<T>> dataListInvoker = PoiInvokerFactory.listInvoker(poiFiledInvokers, ignoreFields);
        PoiInvoker<List<S>> otherDataInvoker = PoiInvokerFactory.listInvoker(poiFiledInvokers, ignoreFields);
        return new PoiIInvokerImpl<T>().generate(sheetName, dataList, dataListInvoker, otherData, otherDataInvoker);
    }


    //
    public static Workbook map2book(String sheetName, List<Map<String, Object>> dataList, Map<String, String> ceTable) {
        PoiInvoker<List<Map<String, Object>>> listPoiInvoker = PoiInvokerFactory.mapInvoker(null, ceTable);
        return new PoiIInvokerImpl<Map<String, Object>>().generate(sheetName, dataList, listPoiInvoker, null, null);
    }

    public static Workbook map2book(String sheetName, List<Map<String, Object>> dataList, Map<String, String> ceTable, String... ignoreFields) {
        PoiInvoker<List<Map<String, Object>>> listPoiInvoker = PoiInvokerFactory.mapInvoker(null, ceTable, ignoreFields);
        return new PoiIInvokerImpl<Map<String, Object>>().generate(sheetName, dataList, listPoiInvoker, null, null);
    }

    public static Workbook map2book(String sheetName,
                                    List<Map<String, Object>> dataList,
                                    Map<String, String> dataListComparisonMap,
                                    List<Map<String, Object>> otherData,
                                    Map<String, String> otherDataComparisonMap) {
        PoiInvoker<List<Map<String, Object>>> dataListInvoker = PoiInvokerFactory.mapInvoker(null, dataListComparisonMap);
        PoiInvoker<List<Map<String, Object>>> otherDataInvoker = PoiInvokerFactory.mapInvoker(null, otherDataComparisonMap);
        return new PoiIInvokerImpl<Map<String, Object>>().generate(sheetName, dataList, dataListInvoker, otherData, otherDataInvoker);
    }

    public static Workbook map2bookWithInvokers(String sheetName,
                                                List<Map<String, Object>> dataList,
                                                Map<String, String> dataListComparisonMap,
                                                List<PoiFiledInvoker> poiFiledInvokers) {
        PoiInvoker<List<Map<String, Object>>> dataListInvoker = PoiInvokerFactory.mapInvoker(poiFiledInvokers, dataListComparisonMap);
        return new PoiIInvokerImpl<Map<String, Object>>().generate(sheetName, dataList, dataListInvoker, null, null);
    }

    public static Workbook map2book(String sheetName,
                                    List<Map<String, Object>> dataList,
                                    Map<String, String> dataListComparisonMap,
                                    List<Map<String, Object>> otherData,
                                    Map<String, String> otherDataComparisonMap,
                                    String... ignoreFields) {
        PoiInvoker<List<Map<String, Object>>> dataListInvoker = PoiInvokerFactory.mapInvoker(null, dataListComparisonMap, ignoreFields);
        PoiInvoker<List<Map<String, Object>>> otherDataInvoker = PoiInvokerFactory.mapInvoker(null, otherDataComparisonMap, ignoreFields);
        return new PoiIInvokerImpl<Map<String, Object>>().generate(sheetName, dataList, dataListInvoker, otherData, otherDataInvoker);
    }

    public static <S, T> Workbook map2book(String sheetName,
                                           List<Map<String, Object>> dataList,
                                           Map<String, String> dataListComparisonMap,
                                           List<Map<String, Object>> otherData,
                                           Map<String, String> otherDataComparisonMap,
                                           List<PoiFiledInvoker> poiFiledInvokers,
                                           String... ignoreFields) {
        PoiInvoker<List<Map<String, Object>>> dataListInvoker = PoiInvokerFactory.mapInvoker(poiFiledInvokers, dataListComparisonMap, ignoreFields);
        PoiInvoker<List<Map<String, Object>>> otherDataInvoker = PoiInvokerFactory.mapInvoker(poiFiledInvokers, otherDataComparisonMap, ignoreFields);
        return new PoiIInvokerImpl<Map<String, Object>>().generate(sheetName, dataList, dataListInvoker, otherData, otherDataInvoker);
    }


    /**
     * workbook转换为java对象,默认忽略第一列
     *
     * @param workbook    workbook对象
     * @param targetClass 目标类型
     * @return List<T>转换后的对象
     */
    public static <T> List<T> book2listImpl(Workbook workbook, Class<T> targetClass) {
        return PoiBookParser.book2listImpl(workbook, targetClass);
    }

    /**
     * workbook转换为java对象,默认忽略第一列
     *
     * @param workbook    workbook对象
     * @param targetClass 目标类型
     * @param ignores 忽略字段属性（匹配字段名）
     * @return List<T>转换后的对象
     */
    public static <T> List<T> book2listImpl(Workbook workbook, Class<T> targetClass,String...ignores) {
        return PoiBookParser.book2listImpl(workbook, targetClass,ignores);
    }

    /**
     * 写出book到response流,通过try-resource自动关闭流
     *
     * @param fileName 文件名
     * @param book     {@link Workbook}对象
     * @param response {@link HttpServletResponse}对象
     * @throws UnsupportedEncodingException
     */
    public static void writeBook(String fileName, Workbook book, HttpServletResponse response) throws UnsupportedEncodingException {
        //处理文件名，下载提示头
        String encodeName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        encodeName = String.format("%s.xlsx", encodeName);
        //下载文件必须的header设置
        String contentDisposition = String.format("attachment;filename=%s", encodeName);
        response.addHeader("Content-Disposition", contentDisposition);
        response.setContentType("application/vnd.ms-excel");
        //写出流
        try (OutputStream out = response.getOutputStream()) {
            book.write(out);
            response.flushBuffer();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
