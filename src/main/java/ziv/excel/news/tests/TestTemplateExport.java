package ziv.excel.news.tests;

import org.apache.poi.ss.usermodel.Workbook;
import ziv.excel.news.PoiUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;

public class TestTemplateExport {
    public static void main(String[] args) throws IOException {
//        exportByClass();
        exportByMap();
    }

    private static void exportByMap() {
        //这个是sheet的第一行表头，key是字段名 value是释义
        Map<String, String> data = new TreeMap<>();
        data.put("test", "测试");
        data.put("test1", "测试1");
        data.put("test2", "测试2");
        data.put("test3", "测试3");
        data.put("test4", "测试4");

        //写出且关流
        try (Workbook sheets = PoiUtil.generateTemplate("测试", data)) {
            String fileName = String.format("D://%s.xlsx", "exportByMap");
            sheets.write(new FileOutputStream(new File(fileName)));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void exportByClass() {
        //写出且关流
        try (Workbook sheets = PoiUtil.generateTemplate("测试", Test.class)) {
            String fileName = String.format("D://%s.xlsx", "exportByClass");
            sheets.write(new FileOutputStream(new File(fileName)));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
