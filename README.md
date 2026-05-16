# shop-order-exporter

一个完整可运行的 Java 8+ Maven 项目，用于本地读取大型订单 CSV 文件，并根据指定配置文件中的「店铺名称 + 商品 ID」关系，按店铺分别导出 CSV 或 XLSX 文件。

> 运行方式以 IDEA 直接点击 `ShopOrderExporter.main()` 为主，不需要打包成 jar 后运行。

## 项目结构

```text
shop-order-exporter/
├── pom.xml
├── README.md
└── src/
    └── main/
        └── java/
            └── com/
                └── example/
                    └── exporter/
                        └── ShopOrderExporter.java
```

## Maven 依赖

项目使用以下核心依赖：

- `org.apache.commons:commons-csv`：流式读取和写入 CSV，兼容字段中包含逗号、引号、换行等场景。
- `org.apache.poi:poi-ooxml`：通过 `SXSSFWorkbook` 流式写入 XLSX，降低大文件导出时的内存占用。

完整 Maven 配置见 [`pom.xml`](pom.xml)。`pom.xml` 只保留编译插件和依赖声明，方便 IDEA 直接导入并运行主程序。

## 输入 CSV 表头

输入文件需要包含以下字段：

```text
支付时间
订单号
商品名称
商品id
招商团长昵称
招商duoid
推手昵称
推手duoid
订单状态
订单金额(元)
招商佣金
招商收入(元)
预估软件服务费(元)
```

程序最终只导出：

```text
支付时间,订单号,商品名称,商品id,订单状态,订单金额(元)
```

## 店铺商品关系配置文件

店铺名称和商品 ID 不需要再写死在 Java 代码里。程序会从固定路径读取配置文件：

```java
private static final String SHOP_GOODS_FILE = "/Users/test/Desktop/shop_goods.txt";
```

文件格式：一行一组数据，店铺名称和商品 ID 之间用空格分隔。

```text
aaa 111
aaa 112
aaa 113
bbb 114
ccc 221
ccc 225
```

说明：

1. 同一个店铺可以配置多行商品 ID。
2. 空行会自动忽略。
3. 以 `#` 开头的注释行会自动忽略。
4. 同一个商品 ID 不允许配置到多个店铺，否则程序会直接报错，避免导出归属不明确。

## 修改固定路径和输出格式

打开 `src/main/java/com/example/exporter/ShopOrderExporter.java`，按实际情况修改这几个常量：

```java
private static final String SOURCE_CSV = "/Users/test/Desktop/order.csv";
private static final String SHOP_GOODS_FILE = "/Users/test/Desktop/shop_goods.txt";
private static final String OUTPUT_DIR = "/Users/test/Desktop/output/";
private static final boolean EXPORT_XLSX = true;
```

`EXPORT_XLSX` 规则：

- `true`：按店铺输出 `aaa.xlsx`、`bbb.xlsx`、`ccc.xlsx`
- `false`：按店铺输出 `aaa.csv`、`bbb.csv`、`ccc.csv`

## 关键逻辑说明

1. 启动后直接读取 `SOURCE_CSV`，无需命令行输入。
2. 启动时先读取 `SHOP_GOODS_FILE`，自动组装成 `Map<String, Set<String>>`。
3. 使用 UTF-8 编码读取订单 CSV 和店铺商品关系文件。
4. 使用 Apache Commons CSV 解析订单 CSV，支持字段中带逗号、引号、换行的标准 CSV。
5. 先把店铺商品配置转换为 `商品 ID -> 店铺名称` 的反向索引，避免每一行订单都遍历所有店铺，提高大文件处理性能。
6. 逐行读取 CSV，不把输入订单文件整体加载到内存。
7. 仅当 `订单状态 = 已成团` 且 `商品id` 命中配置时才导出。
8. 程序启动时为每个店铺动态创建一个输出 writer。
9. 输出目录不存在时会自动创建。

## 如何在 IDEA 中运行

1. 用 IDEA 打开项目根目录。
2. 等待 IDEA 根据 `pom.xml` 下载并加载 Maven 依赖。
3. 修改 `ShopOrderExporter.java` 中的 `SOURCE_CSV`、`SHOP_GOODS_FILE`、`OUTPUT_DIR`、`EXPORT_XLSX`。
4. 准备好订单 CSV 文件和店铺商品关系文件。
5. 直接点击 `ShopOrderExporter` 类中 `main` 方法左侧的运行按钮。
6. 运行完成后，程序会在 `OUTPUT_DIR` 下生成按店铺拆分的文件。

## 可选命令行编译检查

如果本地 Maven 环境可用，也可以执行：

```bash
mvn clean compile
```

但日常使用不需要打包，直接在 IDEA 运行主程序即可。
