package ziv.excel.news.tests;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import ziv.excel.news.PoiConfig;
import ziv.excel.news.PoiUtil;
import ziv.excel.news.invoker.PoiFiledInvoker;
import ziv.excel.news.invoker.fieldInvoker.PoiDefaultDateFieldInvoker;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class TestListExport {


    public static void main(String[] args) throws IOException {
        //循环遍历次数
        int allRecords = 100000;
        //这个是数据行 key是字段名 value是对应的值
        List<Test> lists = new LinkedList<>();
        //遍历插入
        for (int i = 0; i < allRecords; i++) {
            Test test = new Test();
            test.setAge(i);
            test.setName(String.format("刘%s号机器人", i));
            test.setMoney(Math.random());
            test.setBirthday(new Date());
            lists.add(test);
        }
        //这个可以有多个
        List<PoiFiledInvoker> filedInvokers = new ArrayList<>();
        //添加了一个默认的时间字段类型处理器
        filedInvokers.add(new PoiDefaultDateFieldInvoker());

        System.out.println("开始导出...");
        long l = System.currentTimeMillis();
        //写出
        try (Workbook sheets = PoiUtil.list2bookWithInvokers("测试", lists, filedInvokers)) {
            sheets.write(new FileOutputStream(new File("D://TestListExport.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }

        //关流
        System.out.println(String.format("总数据有%s条", allRecords));
        System.out.println(String.format("总共耗时%sms", System.currentTimeMillis() - l));

        //导入
        PoiConfig.IMPORT_VALID_WHOLE_EXCEL = false;
        Workbook sheets = WorkbookFactory.create(new File("D://TestListExport.xlsx"));
        List<Test> tests = PoiUtil.book2listImpl(sheets, Test.class);
        for (Test test : tests) {
            System.out.println(test);
        }
    }
}
