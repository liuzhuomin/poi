package ziv.excel.news.invoker;

/**
 * 处理特殊字段用的
 *
 * @author liuliuliu
 * @since 2021/1/20
 */
public interface PoiFiledInvoker {

	/**
	 * 字段是否匹配
	 *
	 * @return true匹配，false不
	 */
	public boolean match(Object filed);

	/**
	 * 处理字段，获取处理后的字段，可以返回null
	 *
	 * @param obj 字段对应的java获取到的对象上的值
	 * @return 处理过后的值，任意类型
	 */
	public Object getFiledVar(Object obj);
}
