package ziv.excel.news.tests;

import org.apache.poi.ss.usermodel.Workbook;
import ziv.excel.news.PoiUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class TestMapExport {
    public static void main(String[] args) throws IOException {

        //这个是sheet的第一行表头，key是字段名 value是释义
        Map<String, String> config = new TreeMap<>();
        config.put("test", "测试");
        config.put("test1", "测试1");
        config.put("test2", "测试2");
        config.put("test3", "测试3");
        config.put("test4", "测试4");
        //循环遍历次数
        int allRecords = 100000;
        //这个是数据行 key是字段名 value是对应的值
        List<Map<String, Object>> maps = new LinkedList<>();

        for (int i = 0; i < allRecords; i++) {
            Map<String, Object> test1 = new HashMap<>(5);
            test1.put("test", "test1的数据" + i);
            test1.put("test1", "test11的数据" + i);
            test1.put("test2", "test12的数据" + i);
            test1.put("test3", "test13的数据" + i);
            test1.put("test4", "test14的数据" + i);
            maps.add(test1);
        }
        System.out.println("开始导出...");
        long l = System.currentTimeMillis();
        //写出
        try (Workbook sheets = PoiUtil.map2book("测试", maps, config)) {
            sheets.write(new FileOutputStream(new File("D://ziv.excel.news.tests.TestMapExport.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(String.format("总数据有%s条", allRecords));
        System.out.println(String.format("总共耗时%sms", System.currentTimeMillis() - l));
    }
}
