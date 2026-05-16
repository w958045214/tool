package com.example.exporter;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 本地大 CSV 订单文件导出工具。
 *
 * <p>功能：</p>
 * <ul>
 *     <li>从指定配置文件读取「店铺名称 -> 商品 ID 集合」，并据此筛选订单。</li>
 *     <li>仅导出订单状态为「已成团」的数据。</li>
 *     <li>按店铺动态生成 CSV 或 XLSX 文件。</li>
 *     <li>流式读取输入 CSV，避免一次性加载大文件到内存。</li>
 * </ul>
 */
public class ShopOrderExporter {

    /** 输入 CSV 文件路径：按实际环境修改。 */
    private static final String SOURCE_CSV = "/Users/test/Desktop/order.csv";

    /** 店铺商品关系文件路径：一行一组数据，格式为「店铺名称 空格 商品ID」。 */
    private static final String SHOP_GOODS_FILE = "/Users/test/Desktop/shop_goods.txt";

    /** 输出目录：按实际环境修改。 */
    private static final String OUTPUT_DIR = "/Users/test/Desktop/output/";

    /** true 输出 xlsx；false 输出 csv。 */
    private static final boolean EXPORT_XLSX = true;

    /** 输入 CSV 中用于筛选订单状态的固定值。 */
    private static final String TARGET_ORDER_STATUS = "已成团";

    /** 输入 CSV 表头字段。 */
    private static final String COL_PAY_TIME = "支付时间";
    private static final String COL_ORDER_NO = "订单号";
    private static final String COL_GOODS_NAME = "商品名称";
    private static final String COL_GOODS_ID = "商品id";
    private static final String COL_ORDER_STATUS = "订单状态";
    private static final String COL_ORDER_AMOUNT = "订单金额(元)";

    /** 最终输出表头与字段顺序。 */
    private static final List<String> EXPORT_HEADERS = Collections.unmodifiableList(Arrays.asList(
            COL_PAY_TIME,
            COL_ORDER_NO,
            COL_GOODS_NAME,
            COL_GOODS_ID,
            COL_ORDER_STATUS,
            COL_ORDER_AMOUNT
    ));

    public static void main(String[] args) throws IOException {
        Map<String, Set<String>> shopGoodsMap = loadShopGoodsMapFromFile();
        new ShopOrderExporter().export(shopGoodsMap);
    }

    /**
     * 从指定文件读取店铺与商品 ID 的关系。
     *
     * <p>文件格式：一行一组数据，店铺名称和商品 ID 之间使用一个或多个空白字符分隔，例如：</p>
     * <pre>
     * aaa 111
     * aaa 112
     * bbb 114
     * ccc 221
     * </pre>
     *
     * <p>空行和以 # 开头的注释行会被忽略。</p>
     */
    private static Map<String, Set<String>> loadShopGoodsMapFromFile() throws IOException {
        Path configPath = Paths.get(SHOP_GOODS_FILE);
        if (!Files.exists(configPath)) {
            throw new IllegalArgumentException("店铺商品关系文件不存在：" + configPath.toAbsolutePath());
        }

        Map<String, Set<String>> shopGoodsMap = new LinkedHashMap<String, Set<String>>();
        try (BufferedReader reader = Files.newBufferedReader(configPath, StandardCharsets.UTF_8)) {
            String line;
            int lineNo = 0;
            while ((line = reader.readLine()) != null) {
                lineNo++;
                line = line.trim();
                if (line.length() == 0 || line.startsWith("#")) {
                    continue;
                }

                String[] parts = line.split("\\s+");
                if (parts.length != 2 || parts[0].trim().length() == 0 || parts[1].trim().length() == 0) {
                    throw new IllegalArgumentException("店铺商品关系文件第 " + lineNo + " 行格式错误，应为：店铺名称 空格 商品ID");
                }

                String shopName = parts[0].trim();
                String goodsId = parts[1].trim();
                Set<String> goodsIds = shopGoodsMap.get(shopName);
                if (goodsIds == null) {
                    goodsIds = new LinkedHashSet<String>();
                    shopGoodsMap.put(shopName, goodsIds);
                }
                goodsIds.add(goodsId);
            }
        }

        if (shopGoodsMap.isEmpty()) {
            throw new IllegalArgumentException("店铺商品关系文件为空或没有有效配置：" + configPath.toAbsolutePath());
        }
        return shopGoodsMap;
    }

    /**
     * 执行导出。
     */
    public void export(Map<String, Set<String>> shopGoodsMap) throws IOException {
        if (shopGoodsMap == null || shopGoodsMap.isEmpty()) {
            throw new IllegalArgumentException("shopGoodsMap 不能为空");
        }

        Path outputDir = Paths.get(OUTPUT_DIR);
        Files.createDirectories(outputDir);

        Map<String, String> goodsIdToShop = buildGoodsIdToShopIndex(shopGoodsMap);
        Map<String, OrderWriter> writers = createWriters(shopGoodsMap.keySet(), outputDir);

        long totalRows = 0L;
        long exportedRows = 0L;

        try {
            CSVFormat csvFormat = CSVFormat.DEFAULT
                    .withFirstRecordAsHeader()
                    .withIgnoreEmptyLines()
                    .withTrim();

            try (Reader reader = new BufferedReader(new InputStreamReader(
                    new FileInputStream(SOURCE_CSV), StandardCharsets.UTF_8), 1024 * 1024);
                 CSVParser parser = csvFormat.parse(reader)) {

                validateRequiredHeaders(parser.getHeaderMap());

                for (CSVRecord record : parser) {
                    totalRows++;

                    if (!TARGET_ORDER_STATUS.equals(get(record, COL_ORDER_STATUS))) {
                        continue;
                    }

                    String goodsId = get(record, COL_GOODS_ID);
                    String shopName = goodsIdToShop.get(goodsId);
                    if (shopName == null) {
                        continue;
                    }

                    writers.get(shopName).write(toExportRow(record));
                    exportedRows++;
                }
            }
        } finally {
            closeWriters(writers);
        }

        System.out.println("处理完成：读取 " + totalRows + " 行，导出 " + exportedRows + " 行，输出目录：" + outputDir.toAbsolutePath());
    }

    /**
     * 将商品 ID 反向索引到店铺，避免每行 CSV 遍历所有店铺，提升大文件筛选性能。
     */
    private Map<String, String> buildGoodsIdToShopIndex(Map<String, Set<String>> shopGoodsMap) {
        Map<String, String> goodsIdToShop = new HashMap<String, String>();
        for (Map.Entry<String, Set<String>> entry : shopGoodsMap.entrySet()) {
            String shopName = entry.getKey();
            for (String goodsId : entry.getValue()) {
                String oldShop = goodsIdToShop.put(goodsId, shopName);
                if (oldShop != null && !oldShop.equals(shopName)) {
                    throw new IllegalArgumentException("商品 ID " + goodsId + " 同时配置在店铺 " + oldShop + " 和 " + shopName + " 中，请修正配置");
                }
            }
        }
        return goodsIdToShop;
    }

    private Map<String, OrderWriter> createWriters(Set<String> shopNames, Path outputDir) throws IOException {
        Map<String, OrderWriter> writers = new LinkedHashMap<String, OrderWriter>();
        for (String shopName : shopNames) {
            Path outputFile = outputDir.resolve(sanitizeFileName(shopName) + (EXPORT_XLSX ? ".xlsx" : ".csv"));
            OrderWriter writer = EXPORT_XLSX ? new XlsxOrderWriter(outputFile) : new CsvOrderWriter(outputFile);
            writer.writeHeader(EXPORT_HEADERS);
            writers.put(shopName, writer);
        }
        return writers;
    }

    private void validateRequiredHeaders(Map<String, Integer> headerMap) {
        List<String> missingHeaders = new ArrayList<String>();
        for (String header : EXPORT_HEADERS) {
            if (!headerMap.containsKey(header)) {
                missingHeaders.add(header);
            }
        }
        if (!missingHeaders.isEmpty()) {
            throw new IllegalArgumentException("输入 CSV 缺少必要表头：" + missingHeaders);
        }
    }

    private List<String> toExportRow(CSVRecord record) {
        List<String> row = new ArrayList<String>(EXPORT_HEADERS.size());
        for (String header : EXPORT_HEADERS) {
            row.add(get(record, header));
        }
        return row;
    }

    private static String get(CSVRecord record, String header) {
        String value = record.get(header);
        return value == null ? "" : value.trim();
    }

    private static String sanitizeFileName(String fileName) {
        return fileName.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    private static void closeWriters(Map<String, OrderWriter> writers) throws IOException {
        IOException firstException = null;
        for (OrderWriter writer : writers.values()) {
            try {
                writer.close();
            } catch (IOException e) {
                if (firstException == null) {
                    firstException = e;
                } else {
                    firstException.addSuppressed(e);
                }
            }
        }
        if (firstException != null) {
            throw firstException;
        }
    }

    /** 订单输出抽象，便于 CSV / XLSX 两种格式切换。 */
    private interface OrderWriter extends Closeable {
        void writeHeader(List<String> headers) throws IOException;

        void write(List<String> values) throws IOException;
    }

    /** CSV 输出实现：使用 Apache Commons CSV 正确处理逗号、引号、换行等特殊字符。 */
    private static class CsvOrderWriter implements OrderWriter {
        private final CSVPrinter printer;

        CsvOrderWriter(Path outputFile) throws IOException {
            Writer writer = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream(outputFile.toFile()), StandardCharsets.UTF_8), 1024 * 1024);
            this.printer = new CSVPrinter(writer, CSVFormat.DEFAULT);
        }

        @Override
        public void writeHeader(List<String> headers) throws IOException {
            printer.printRecord(headers);
        }

        @Override
        public void write(List<String> values) throws IOException {
            printer.printRecord(values);
        }

        @Override
        public void close() throws IOException {
            printer.close(true);
        }
    }

    /** XLSX 输出实现：使用 SXSSFWorkbook 流式写入，适合较大的导出结果。 */
    private static class XlsxOrderWriter implements OrderWriter {
        private static final int ROW_ACCESS_WINDOW_SIZE = 500;

        private final Path outputFile;
        private final SXSSFWorkbook workbook;
        private final Sheet sheet;
        private final CellStyle headerStyle;
        private int rowIndex = 0;

        XlsxOrderWriter(Path outputFile) {
            this.outputFile = outputFile;
            this.workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
            this.workbook.setCompressTempFiles(true);
            this.sheet = workbook.createSheet("orders");
            this.headerStyle = createHeaderStyle(workbook);
        }

        @Override
        public void writeHeader(List<String> headers) {
            Row row = sheet.createRow(rowIndex++);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(headers.get(i));
                cell.setCellStyle(headerStyle);
            }
        }

        @Override
        public void write(List<String> values) {
            Row row = sheet.createRow(rowIndex++);
            for (int i = 0; i < values.size(); i++) {
                row.createCell(i).setCellValue(values.get(i));
            }
        }

        @Override
        public void close() throws IOException {
            for (int i = 0; i < EXPORT_HEADERS.size(); i++) {
                sheet.setColumnWidth(i, 20 * 256);
            }
            try (FileOutputStream outputStream = new FileOutputStream(outputFile.toFile())) {
                workbook.write(outputStream);
            } finally {
                workbook.dispose();
                workbook.close();
            }
        }

        private static CellStyle createHeaderStyle(SXSSFWorkbook workbook) {
            Font font = workbook.createFont();
            font.setBold(true);

            CellStyle style = workbook.createCellStyle();
            style.setFont(font);
            return style;
        }
    }
}
