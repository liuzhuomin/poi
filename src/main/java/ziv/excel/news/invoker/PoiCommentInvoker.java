package ziv.excel.news.invoker;

/**
 * 负责解析用的
 *
 * @author liuliuliu
 * @since 2021/1/20
 */
public interface PoiCommentInvoker {

	/**
	 * 获取当前单个字段的中文注释
	 *
	 * @param sourceField 字段
	 * @return 中文注释或者英文注释
	 */
	String getComment(Object sourceField);
}
