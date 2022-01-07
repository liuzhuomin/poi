package ziv.excel.news.invoker;

import java.util.List;
import java.util.Map;

/**
 * poi执行器
 *
 * @param <T>
 * @author liuliuliu
 * @since 20201/1/20
 */
public interface PoiInvoker<T> {

	/**
	 * 获取表头
	 *
	 * @param source 源对象
	 * @return 顺序map
	 */
	Map<String, String> getHeaders(T source);

	/**
	 * 获取总行
	 *
	 * @param source 源对象
	 * @return 顺序map
	 */
	List<Map<String, Object>> getRows(T source);

	/**
	 * 获取poi样式对象
	 *
	 * @return 样式对象
	 */
	PoiStyleInvoker getPoiStyleInvoker();

}
