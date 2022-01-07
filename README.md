## 一款方便快捷的excel导入导出工具

### 声明
- 不听废话的直接看后面代码就行。

### 导出
- 导出支持map转excel和list转execl,sheet页数据满了自动创建第二页，支持在sheet页码最后增加统计数据（有的业务需要）
暂时不支持单元格等复杂模式；


- 导出支持自己配置解析，包含`注释获取规则`、`cell列样式`、`支持统计列`,`支持字段特殊处理`

#### map转excel

- 核心code
```java
//参数1是sheet名;参数2是一个装载着Map的List集合，参数三是对map每个字段做释义的普通map
Workbook sheets = PoiUtil.map2book("测试", maps, config)
```
- 复杂案例
```java
package ziv.excel.news;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ziv.excel.news.tests.TestMapExport {
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
        //map to book对象
        //写出
        try (Workbook sheets = PoiUtil.map2book("测试", maps, config)) {
            sheets.write(new FileOutputStream(new File("D://ziv.excel.news.tests.TestMapExport.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }
        //关流
        System.out.println(String.format("总数据有%s条", allRecords));
        System.out.println(String.format("总共耗时%sms", System.currentTimeMillis() - l));
    }
}

```

### List转excel

- 核心code
```java
Workbook sheets = PoiUtil.list2book("测试", lists);
```
- 复杂案例
```java
package ziv.excel.news.tests;

public class ziv.excel.news.tests.TestListExport {

    public static void main(String[] args) {
        //循环遍历次数
        int allRecords = 100000;
        //这个是数据行 key是字段名 value是对应的值
        List<ziv.excel.news.tests.Test> lists = new LinkedList<>();
        //遍历插入
        for (int i = 0; i < allRecords; i++) {
            ziv.excel.news.tests.Test test = new ziv.excel.news.tests.Test();
            test.setAge(i);
            test.setName(String.format("刘%s号机器人", i));
            test.setMoeny(Math.random());
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
            sheets.write(new FileOutputStream(new File("D://ziv.excel.news.tests.TestListExport.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        //关流
        System.out.println(String.format("总数据有%s条", allRecords));
        System.out.println(String.format("总共耗时%sms", System.currentTimeMillis() - l));
    }
}

```

### 模板生成
```java
    public static void main(String[] args) throws IOException {
        System.out.println("开始导出...");
        long l = System.currentTimeMillis();
        //写出且关流
        try (Workbook sheets = PoiUtil.generateTemplate("测试", ziv.excel.news.tests.Test.class)) {
            String fileName = String.format("D://%s.xlsx", ziv.excel.news.tests.TestTemplateExport.class.getSimpleName());
            sheets.write(new FileOutputStream(new File(fileName)));
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(String.format("总共耗时%sms", System.currentTimeMillis() - l));
    }
```


### 注释获取规则

- 基础类为`ziv.excel.news.invoker.PoiCommentInvoker`

