# shop-order-exporter

一个完整可运行的 Java 8+ Maven 项目，用于本地读取大型订单 CSV 文件，并根据代码中预设的「店铺名称 + 商品 ID」配置，按店铺分别导出 CSV 或 XLSX 文件。

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

完整 Maven 配置见 [`pom.xml`](pom.xml)。

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

## 修改固定配置

打开 `src/main/java/com/example/exporter/ShopOrderExporter.java`，按实际情况修改输入文件、输出目录、输出格式开关：

```java
private static final String SOURCE_CSV = "/Users/test/Desktop/order.csv";
private static final String OUTPUT_DIR = "/Users/test/Desktop/output/";
private static final boolean EXPORT_XLSX = true;
```

修改店铺与商品 ID 配置：

```java
shopGoodsMap.put("aaa", setOf("111", "112", "113"));
shopGoodsMap.put("bbb", setOf("114"));
shopGoodsMap.put("ccc", setOf("221", "225"));
```

## 关键逻辑说明

1. 启动后直接读取 `SOURCE_CSV`，无需命令行输入。
2. 使用 UTF-8 编码和 Apache Commons CSV 解析输入文件，支持字段中带逗号的标准 CSV。
3. 先把配置转换为 `商品 ID -> 店铺名称` 的反向索引，避免每一行都遍历所有店铺，提高大文件处理性能。
4. 逐行读取 CSV，不把输入文件整体加载到内存。
5. 仅当 `订单状态 = 已成团` 且 `商品id` 命中配置时才导出。
6. 程序启动时为每个店铺动态创建一个输出 writer。
7. `EXPORT_XLSX = true` 时输出 `aaa.xlsx`、`bbb.xlsx`、`ccc.xlsx`；`false` 时输出 `aaa.csv`、`bbb.csv`、`ccc.csv`。
8. 输出目录不存在时会自动创建。

## 如何运行

### 1. 编译

```bash
mvn clean package
```

### 2. 运行

```bash
java -jar target/shop-order-exporter-1.0.0.jar
```

运行完成后，程序会在 `OUTPUT_DIR` 下生成按店铺拆分的文件。
