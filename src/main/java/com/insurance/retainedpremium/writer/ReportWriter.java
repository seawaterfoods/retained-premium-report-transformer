package com.insurance.retainedpremium.writer;

import com.insurance.retainedpremium.config.InsuranceMappingService;
import com.insurance.retainedpremium.config.InsuranceMappingService.CategoryDef;
import com.insurance.retainedpremium.config.InsuranceMappingService.GuishuEntry;
import com.insurance.retainedpremium.model.CompanyData;
import com.insurance.retainedpremium.model.QuarterData;
import com.insurance.retainedpremium.service.DataTransformerService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static com.insurance.retainedpremium.constant.InsuranceConstants.*;
import static com.insurance.retainedpremium.writer.ExcelStyleHelper.*;

/**
 * 程式化產生報表 Excel — 不依賴模板檔案
 */
@Component
public class ReportWriter {

    private static final Logger log = LoggerFactory.getLogger(ReportWriter.class);

    private final DataTransformerService dataTransformerService;
    private final InsuranceMappingService mapping;

    public ReportWriter(DataTransformerService dataTransformerService, InsuranceMappingService mapping) {
        this.dataTransformerService = dataTransformerService;
        this.mapping = mapping;
    }

    /**
     * 產生報表 Excel 檔案
     *
     * @param outputPath     輸出檔案路徑
     * @param quarterDataMap 季度資料 (quarter → QuarterData)
     * @param lastYearData   去年同期資料 (companyCode → total)
     * @param year           民國年
     * @param maxQuarter     最大季度
     */
    public void writeReport(String outputPath,
                            Map<Integer, QuarterData> quarterDataMap,
                            Map<String, Double> lastYearData,
                            int year, int maxQuarter) throws IOException {

        Files.createDirectories(Path.of(outputPath).getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            ExcelStyleHelper styles = new ExcelStyleHelper(workbook);

            String sheet1Name = year + "自留(季)";
            String sheet2Name = year + "自留總累";

            XSSFSheet sheet1 = workbook.createSheet(sheet1Name);
            XSSFSheet sheet2 = workbook.createSheet(sheet2Name);
            XSSFSheet sheet3 = workbook.createSheet("歸屬");

            // 建立公司排序清單 (所有季度的聯集，依代號排序)
            List<String[]> companyList = buildCompanyList(quarterDataMap);

            // Sheet1: 季度明細
            Map<Integer, int[]> s1QuarterRowMap = writeSheet1(sheet1, styles, quarterDataMap,
                    companyList, year);

            // Sheet2: 累計總表
            writeSheet2(sheet2, sheet1Name, styles, quarterDataMap, lastYearData,
                    companyList, year, maxQuarter, s1QuarterRowMap);

            // Sheet3: 歸屬表
            writeGuishuSheet(sheet3, styles);

            workbook.setForceFormulaRecalculation(true);

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
            }

            log.info("Report written successfully to {}", outputPath);
        }
    }

    /**
     * 從所有季度資料建立排序的公司清單 [code, name]
     */
    private List<String[]> buildCompanyList(Map<Integer, QuarterData> quarterDataMap) {
        Map<String, String> codeToName = new LinkedHashMap<>();
        for (QuarterData qd : quarterDataMap.values()) {
            for (CompanyData cd : qd.companies().values()) {
                codeToName.putIfAbsent(cd.companyCode(), cd.companyName());
            }
        }
        List<String[]> list = new ArrayList<>();
        codeToName.entrySet().stream()
                .sorted(Comparator.comparingInt(e -> Integer.parseInt(e.getKey())))
                .forEach(e -> list.add(new String[]{e.getKey(), e.getValue()}));
        return list;
    }

    // ====================================================================
    // Sheet1: 季度明細表
    // ====================================================================

    /**
     * @return quarter → [dataStartRow(0-based), dataEndRow(0-based)] for cross-sheet references
     */
    private Map<Integer, int[]> writeSheet1(XSSFSheet sheet, ExcelStyleHelper styles,
                                             Map<Integer, QuarterData> quarterDataMap,
                                             List<String[]> companyList, int year) {
        // --- Column widths ---
        sheet.setColumnWidth(0, charWidth(4.22));   // A: 代號
        sheet.setColumnWidth(1, charWidth(10.78));  // B: 月份
        sheet.setColumnWidth(2, charWidth(13.89));  // C: 公司別
        for (int c = S1_COL_DATA_START - 1; c <= mapping.getS1ColDataEnd() - 1; c++) {
            sheet.setColumnWidth(c, charWidth(13.5));
        }
        sheet.setColumnWidth(mapping.getS1ColTotal() - 1, charWidth(16.22)); // 合計

        // --- Header Row 0: insurance codes ---
        Row row0 = sheet.createRow(S1_HEADER_ROW_CODES);
        row0.setHeightInPoints(ExcelStyleHelper.S1_ROW0_HEIGHT);
        createCell(row0, 0, (String) null, styles.getS1CodeValueStyle());
        createCell(row0, 1, (String) null, styles.getS1CodeValueBoldStyle());
        createCell(row0, 2, "險種代號", styles.getS1CodeHeaderStyle());
        for (String code : mapping.getInsuranceCodes()) {
            int col = mapping.getS1CodeToCol().get(code) - 1;
            createCell(row0, col, code, styles.getS1CodeValueStyle());
        }

        // --- Header Row 1: insurance names + labels ---
        Row row1 = sheet.createRow(S1_HEADER_ROW_NAMES);
        row1.setHeightInPoints(ExcelStyleHelper.S1_ROW1_HEIGHT);
        createCell(row1, 0, "代號", styles.getS1LabelHeaderStyle());
        createCell(row1, 1, "月份", styles.getS1MonthHeaderStyle());
        createCell(row1, 2, "公司別/險種", styles.getS1CompanyNameHeaderStyle());
        for (String code : mapping.getInsuranceCodes()) {
            int col = mapping.getS1CodeToCol().get(code) - 1;
            createCell(row1, col, mapping.getCodeToShortName().get(code), styles.getS1NameHeaderStyle());
        }
        createCell(row1, mapping.getS1ColTotal() - 1, "合計", styles.getS1NameHeaderStyle());

        // --- Freeze panes (D2) ---
        sheet.createFreezePane(3, 1);

        // --- Quarter blocks ---
        Map<Integer, int[]> quarterRowMap = new LinkedHashMap<>();
        int currentRow = S1_DATA_START_ROW; // 0-based row 2

        for (int q = 1; q <= 4; q++) {
            QuarterData qd = quarterDataMap.get(q);
            if (qd == null) continue;

            String monthLabel = QUARTER_LABEL_MAP.get(q);
            int blockStart = currentRow;

            // Company data rows
            for (String[] company : companyList) {
                String code = company[0];
                String name = company[1];
                CompanyData cd = qd.companies().get(code);

                Row dataRow = sheet.createRow(currentRow);
                dataRow.setHeightInPoints(ExcelStyleHelper.S1_DATA_ROW_HEIGHT);
                createCell(dataRow, 0, code, styles.getS1CompanyCodeStyle());
                createCell(dataRow, 1, monthLabel, styles.getS1MonthStyle());
                createCell(dataRow, 2, name, styles.getS1CompanyNameStyle());

                // Insurance code values
                for (String insCode : mapping.getInsuranceCodes()) {
                    int col = mapping.getS1CodeToCol().get(insCode) - 1;
                    double value = 0.0;
                    if (cd != null) {
                        value = cd.retainedPremiums().getOrDefault(insCode, 0.0);
                    }
                    createCell(dataRow, col, value, styles.getS1DataStyle());
                }

                // Total SUM formula
                String firstDataCol = colLetter(S1_COL_DATA_START - 1);
                String lastDataCol = colLetter(mapping.getS1ColDataEnd() - 1);
                int excelRow = currentRow + 1;
                createFormulaCell(dataRow, mapping.getS1ColTotal() - 1,
                        "SUM(" + firstDataCol + excelRow + ":" + lastDataCol + excelRow + ")",
                        styles.getS1DataStyle());

                // Hide row if company has no data for this quarter
                if (cd == null) {
                    dataRow.setZeroHeight(true);
                }

                currentRow++;
            }

            int blockEnd = currentRow - 1;
            quarterRowMap.put(q, new int[]{blockStart, blockEnd});

            // Subtotal row
            Row subtotalRow = sheet.createRow(currentRow);
            subtotalRow.setHeightInPoints(ExcelStyleHelper.S1_DATA_ROW_HEIGHT);
            createCell(subtotalRow, 0, "", styles.getS1SubtotalCodeStyle());
            createCell(subtotalRow, 1, monthLabel, styles.getS1SubtotalMonthStyle());
            createCell(subtotalRow, 2, "小計", styles.getS1SubtotalLabelStyle());

            for (int col = S1_COL_DATA_START - 1; col <= mapping.getS1ColTotal() - 1; col++) {
                String colL = colLetter(col);
                String formula = "SUBTOTAL(9," + colL + (blockStart + 1) + ":" + colL + (blockEnd + 1) + ")";
                createFormulaCell(subtotalRow, col, formula, styles.getS1SubtotalDataStyle());
            }

            currentRow++;
        }

        // Auto filter on header row
        if (currentRow > S1_DATA_START_ROW) {
            sheet.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(
                    S1_HEADER_ROW_NAMES, currentRow - 1, 0, mapping.getS1ColTotal() - 1));
        }

        return quarterRowMap;
    }

    // ====================================================================
    // Sheet2: 累計總表
    // ====================================================================

    private void writeSheet2(XSSFSheet sheet, String sheet1Name, ExcelStyleHelper styles,
                             Map<Integer, QuarterData> quarterDataMap,
                             Map<String, Double> lastYearData,
                             List<String[]> companyList,
                             int year, int maxQuarter,
                             Map<Integer, int[]> s1QuarterRowMap) {

        // --- Column widths ---
        sheet.setColumnWidth(0, charWidth(4.89));
        sheet.setColumnWidth(1, charWidth(12.89));
        sheet.setColumnWidth(2, charWidth(17.22));
        for (int c = S2_COL_DATA_START - 1; c <= mapping.getS2ColDataEnd() - 1; c++) {
            sheet.setColumnWidth(c, charWidth(17.56));
        }
        sheet.setColumnWidth(mapping.getS2ColYearTotal() - 1, charWidth(19.22));
        sheet.setColumnWidth(mapping.getS2ColLastYear() - 1, charWidth(13.0));
        sheet.setColumnWidth(mapping.getS2ColGrowth() - 1, charWidth(12.89));

        // --- Row 0: Title ---
        int[] months = QUARTER_MONTH_MAP.get(maxQuarter);
        String title = year + "年度第" + maxQuarter + "季(" + months[0] + "-" + months[1]
                + "月份)各保險公司自留保費統計總表";
        Row titleRow = sheet.createRow(S2_HEADER_ROW_TITLE);
        titleRow.setHeightInPoints(ExcelStyleHelper.S2_ROW0_HEIGHT);
        createCell(titleRow, 0, title, styles.getS2TitleStyle());
        mergeRegion(sheet, 0, 0, 0, mapping.getS2ColYearTotal() - 1);

        // --- Row 1: Unit ---
        Row unitRow = sheet.createRow(S2_HEADER_ROW_UNIT);
        unitRow.setHeightInPoints(ExcelStyleHelper.S2_ROW1_HEIGHT);
        createCell(unitRow, mapping.getS2ColGrowth() - 1, "單位:新台幣元", styles.getS2UnitStyle());

        // --- Row 2: Category group headers ---
        Row groupRow = sheet.createRow(S2_HEADER_ROW_GROUP);
        groupRow.setHeightInPoints(ExcelStyleHelper.S2_ROW2_HEIGHT);
        Map<String, int[]> catGroups = mapping.getS2CategoryGroups();
        for (Map.Entry<String, int[]> entry : catGroups.entrySet()) {
            String groupName = entry.getKey();
            if ("比較".equals(groupName)) continue; // 比較放在最後
            int startCol = entry.getValue()[0] - 1;
            int endCol = entry.getValue()[1] - 1;
            createCell(groupRow, startCol, groupName, styles.getS2GroupHeaderStyle());
            if (startCol != endCol) {
                mergeRegion(sheet, S2_HEADER_ROW_GROUP, S2_HEADER_ROW_GROUP, startCol, endCol);
            }
            for (int c = startCol + 1; c <= endCol; c++) {
                createCell(groupRow, c, "", styles.getS2GroupHeaderStyle());
            }
        }
        createCell(groupRow, mapping.getS2ColGrowth() - 1, "比較", styles.getS2GroupHeaderStyle());

        // --- Row 3: Main headers ---
        Row mainRow = sheet.createRow(S2_HEADER_ROW_MAIN);
        mainRow.setHeightInPoints(ExcelStyleHelper.S2_ROW3_HEIGHT);
        createCell(mainRow, 0, "代號", styles.getS2HeaderCodeStyle());
        createCell(mainRow, 1, "月份", styles.getS2HeaderCompanyStyle());
        createCell(mainRow, 2, "公司別/險種", styles.getS2HeaderCompanyStyle());

        Map<String, CategoryDef> catMapping = mapping.getCategoryMapping();
        List<String> categoryNames = new ArrayList<>(catMapping.keySet());
        for (int i = 0; i < categoryNames.size(); i++) {
            String catName = categoryNames.get(i);
            int col = catMapping.get(catName).columnIndex() - 1;
            createCell(mainRow, col, catName, styles.getS2MainHeaderStyle());
        }
        createCell(mainRow, mapping.getS2ColYearTotal() - 1, year + "年度", styles.getS2MainHeaderStyle());
        createCell(mainRow, mapping.getS2ColLastYear() - 1, (year - 1) + "年度", styles.getS2MainHeaderStyle());
        createCell(mainRow, mapping.getS2ColGrowth() - 1, "成長率", styles.getS2MainHeaderStyle());

        // --- Row 4: Sub headers ---
        Row subRow = sheet.createRow(S2_HEADER_ROW_SUB);
        subRow.setHeightInPoints(ExcelStyleHelper.S2_ROW4_HEIGHT);
        Map<Integer, String> subHeaders = mapping.getS2SubHeaders();
        for (Map.Entry<Integer, String> entry : subHeaders.entrySet()) {
            createCell(subRow, entry.getKey() - 1, entry.getValue(), styles.getS2SubHeaderStyle());
        }
        createCell(subRow, mapping.getS2ColYearTotal() - 1, "合計", styles.getS2SubHeaderStyle());
        createCell(subRow, mapping.getS2ColLastYear() - 1, "合計", styles.getS2SubHeaderStyle());
        createCell(subRow, mapping.getS2ColGrowth() - 1, "%", styles.getS2SubHeaderStyle());

        // Fill empty sub-header cells with style for consistency
        for (int c = 0; c < mapping.getS2ColGrowth(); c++) {
            if (subRow.getCell(c) == null) {
                createCell(subRow, c, "", styles.getS2SubHeaderStyle());
            }
        }

        // Merged cells for headers that span rows 3-4
        mergeRegion(sheet, S2_HEADER_ROW_MAIN, S2_HEADER_ROW_SUB, 0, 0);
        mergeRegion(sheet, S2_HEADER_ROW_MAIN, S2_HEADER_ROW_SUB, 1, 1);
        mergeRegion(sheet, S2_HEADER_ROW_MAIN, S2_HEADER_ROW_SUB, 2, 2);
        // Categories without sub-headers span rows 3-4
        for (String catName : categoryNames) {
            int col = catMapping.get(catName).columnIndex() - 1;
            boolean hasSub = subHeaders.containsKey(col + 1);
            if (!hasSub) {
                mergeRegion(sheet, S2_HEADER_ROW_MAIN, S2_HEADER_ROW_SUB, col, col);
            }
        }
        // Header group merges (e.g. 強制責任險)
        for (int[] merge : mapping.getS2HeaderGroupMerges()) {
            mergeRegion(sheet, S2_HEADER_ROW_MAIN, S2_HEADER_ROW_MAIN, merge[0] - 1, merge[1] - 1);
        }

        // --- Freeze panes ---
        sheet.createFreezePane(3, S2_DATA_START_ROW);

        // --- Quarter data blocks ---
        int currentRow = S2_DATA_START_ROW;
        List<int[]> allQuarterBlocks = new ArrayList<>();

        // Category cross-sheet formula builder
        List<CategoryDef> categories = new ArrayList<>(catMapping.values());

        for (int q = 1; q <= 4; q++) {
            QuarterData qd = quarterDataMap.get(q);
            if (qd == null) continue;

            String monthLabel = QUARTER_LABEL_MAP.get(q);
            int[] s1Block = s1QuarterRowMap.get(q);
            int blockStart = currentRow;

            for (int ci = 0; ci < companyList.size(); ci++) {
                String[] company = companyList.get(ci);
                String code = company[0];
                String name = company[1];
                CompanyData cd = qd.companies().get(code);

                Row dataRow = sheet.createRow(currentRow);
                dataRow.setHeightInPoints(ExcelStyleHelper.S2_DATA_ROW_HEIGHT);
                createCell(dataRow, 0, code, styles.getS2CompanyCodeStyle());
                createCell(dataRow, 1, monthLabel, styles.getS2MonthStyle());
                createCell(dataRow, 2, name, styles.getS2CompanyNameStyle());

                // Cross-sheet formulas for categories (D-S)
                int s1DataRow = s1Block[0] + ci; // 0-based row in Sheet1
                int s1ExcelRow = s1DataRow + 1;  // 1-based for formula
                String quotedSheet1 = "'" + sheet1Name + "'!";

                for (CategoryDef catDef : categories) {
                    int s2Col = catDef.columnIndex() - 1;
                    List<String> codes = catDef.insuranceCodes();

                    if (codes.size() == 1) {
                        int s1Col = mapping.getS1CodeToCol().get(codes.get(0)) - 1;
                        String formula = quotedSheet1 + colLetter(s1Col) + s1ExcelRow;
                        createFormulaCell(dataRow, s2Col, formula, styles.getS2DataStyle());
                    } else {
                        // SUM of multiple columns
                        String firstCol = colLetter(mapping.getS1CodeToCol().get(codes.get(0)) - 1);
                        String lastCol = colLetter(mapping.getS1CodeToCol().get(codes.get(codes.size() - 1)) - 1);
                        // Check if columns are contiguous
                        int first = mapping.getS1CodeToCol().get(codes.get(0));
                        int last = mapping.getS1CodeToCol().get(codes.get(codes.size() - 1));
                        if (last - first + 1 == codes.size()) {
                            String formula = "SUM(" + quotedSheet1 + firstCol + s1ExcelRow
                                    + ":" + lastCol + s1ExcelRow + ")";
                            createFormulaCell(dataRow, s2Col, formula, styles.getS2DataStyle());
                        } else {
                            // Non-contiguous: sum individual cells
                            StringBuilder formula = new StringBuilder();
                            for (int i = 0; i < codes.size(); i++) {
                                if (i > 0) formula.append("+");
                                int col = mapping.getS1CodeToCol().get(codes.get(i)) - 1;
                                formula.append(quotedSheet1).append(colLetter(col)).append(s1ExcelRow);
                            }
                            createFormulaCell(dataRow, s2Col, formula.toString(), styles.getS2DataStyle());
                        }
                    }
                }

                // T: Year total = SUM(D:last-cat)
                int excelRow = currentRow + 1;
                String dCol = colLetter(S2_COL_DATA_START - 1);
                String sCol = colLetter(mapping.getS2ColDataEnd() - 1);
                createFormulaCell(dataRow, mapping.getS2ColYearTotal() - 1,
                        "SUM(" + dCol + excelRow + ":" + sCol + excelRow + ")",
                        styles.getS2DataStyle());

                // U: Last year data
                Double lastYear = lastYearData != null ? lastYearData.get(code) : null;
                if (lastYear != null) {
                    createCell(dataRow, mapping.getS2ColLastYear() - 1, lastYear, styles.getS2DataStyle());
                } else {
                    createCell(dataRow, mapping.getS2ColLastYear() - 1, "", styles.getS2DataStyle());
                    if (cd != null) {
                        log.warn("公司代號 {} 無去年資料，U欄留空", code);
                    }
                }

                // V: Growth rate = IF(U=0,"",((T/U)-1))
                String uRef = colLetter(mapping.getS2ColLastYear() - 1) + excelRow;
                String tRef = colLetter(mapping.getS2ColYearTotal() - 1) + excelRow;
                createFormulaCell(dataRow, mapping.getS2ColGrowth() - 1,
                        "IF(" + uRef + "=0,\"\",(" + tRef + "/" + uRef + ")-1)",
                        styles.getS2PercentStyle());

                // Hide row if no data
                if (cd == null) {
                    dataRow.setZeroHeight(true);
                }

                currentRow++;
            }

            int blockEnd = currentRow - 1;
            allQuarterBlocks.add(new int[]{blockStart, blockEnd});

            // Subtotal row
            Row subtotalRow = sheet.createRow(currentRow);
            subtotalRow.setHeightInPoints(ExcelStyleHelper.S2_DATA_ROW_HEIGHT);
            createCell(subtotalRow, 0, "", styles.getS2SubtotalCodeStyle());
            createCell(subtotalRow, 1, monthLabel, styles.getS2SubtotalMonthStyle());
            createCell(subtotalRow, 2, "小計", styles.getS2SubtotalLabelStyle());

            int excelBlockStart = blockStart + 1;
            int excelBlockEnd = blockEnd + 1;

            for (int col = S2_COL_DATA_START - 1; col <= mapping.getS2ColYearTotal() - 1; col++) {
                String colL = colLetter(col);
                createFormulaCell(subtotalRow, col,
                        "SUBTOTAL(9," + colL + excelBlockStart + ":" + colL + excelBlockEnd + ")",
                        styles.getS2SubtotalDataStyle());
            }
            // U subtotal
            String uColL = colLetter(mapping.getS2ColLastYear() - 1);
            createFormulaCell(subtotalRow, mapping.getS2ColLastYear() - 1,
                    "SUBTOTAL(9," + uColL + excelBlockStart + ":" + uColL + excelBlockEnd + ")",
                    styles.getS2SubtotalDataStyle());
            // V growth for subtotal: IF(U=0,"",((T-U)/U))
            int subtotalExcelRow = currentRow + 1;
            String stU = colLetter(mapping.getS2ColLastYear() - 1) + subtotalExcelRow;
            String stT = colLetter(mapping.getS2ColYearTotal() - 1) + subtotalExcelRow;
            createFormulaCell(subtotalRow, mapping.getS2ColGrowth() - 1,
                    "IF(" + stU + "=0,\"\",(" + stT + "-" + stU + ")/" + stU + ")",
                    styles.getS2SubtotalPctStyle());

            currentRow++;
        }

        // Grand total row
        if (!allQuarterBlocks.isEmpty()) {
            Row grandTotalRow = sheet.createRow(currentRow);
            grandTotalRow.setHeightInPoints(ExcelStyleHelper.S2_DATA_ROW_HEIGHT);
            createCell(grandTotalRow, 0, "", styles.getS2SubtotalCodeStyle());
            createCell(grandTotalRow, 1, "總計", styles.getS2SubtotalMonthStyle());
            createCell(grandTotalRow, 2, "", styles.getS2SubtotalLabelStyle());

            int firstBlock = allQuarterBlocks.get(0)[0] + 1;
            int lastBlock = currentRow; // includes all subtotal rows

            for (int col = S2_COL_DATA_START - 1; col <= mapping.getS2ColYearTotal() - 1; col++) {
                String colL = colLetter(col);
                createFormulaCell(grandTotalRow, col,
                        "SUBTOTAL(9," + colL + firstBlock + ":" + colL + lastBlock + ")",
                        styles.getS2SubtotalDataStyle());
            }
            String uColL = colLetter(mapping.getS2ColLastYear() - 1);
            createFormulaCell(grandTotalRow, mapping.getS2ColLastYear() - 1,
                    "SUBTOTAL(9," + uColL + firstBlock + ":" + uColL + lastBlock + ")",
                    styles.getS2SubtotalDataStyle());

            int gtExcelRow = currentRow + 1;
            String gtU = colLetter(mapping.getS2ColLastYear() - 1) + gtExcelRow;
            String gtT = colLetter(mapping.getS2ColYearTotal() - 1) + gtExcelRow;
            createFormulaCell(grandTotalRow, mapping.getS2ColGrowth() - 1,
                    "IF(" + gtT + "=0,\"\",(" + gtT + "/" + gtU + ")-1)",
                    styles.getS2SubtotalPctStyle());
        }
    }

    // ====================================================================
    // Sheet3: 歸屬表
    // ====================================================================

    private void writeGuishuSheet(XSSFSheet sheet, ExcelStyleHelper styles) {
        sheet.setColumnWidth(0, charWidth(6));
        sheet.setColumnWidth(1, charWidth(12));
        sheet.setColumnWidth(2, charWidth(8));
        sheet.setColumnWidth(3, charWidth(2));
        sheet.setColumnWidth(4, charWidth(22));

        // Header row
        Row header = sheet.createRow(0);
        createCell(header, 0, "類", styles.getGuishuHeaderStyle());
        createCell(header, 1, "歸屬", styles.getGuishuHeaderStyle());
        createCell(header, 2, "代號", styles.getGuishuHeaderStyle());
        createCell(header, 4, "險種", styles.getGuishuHeaderStyle());

        // Data rows
        int rowIdx = 1;
        for (GuishuEntry entry : mapping.getGuishuTable()) {
            Row row = sheet.createRow(rowIdx++);
            if (entry.category() != null) {
                createCell(row, 0, entry.category(), styles.getGuishuCategoryStyle());
            }
            if (entry.group() != null) {
                createCell(row, 1, entry.group(), styles.getGuishuDataStyle());
            }
            createCell(row, 2, entry.code(), styles.getGuishuCodeStyle());
            createCell(row, 4, entry.name(), styles.getGuishuDataStyle());
        }
    }

    // ====================================================================
    // 工具方法
    // ====================================================================

    /** 將字元寬度轉換為 POI column width 單位 */
    private static int charWidth(double chars) {
        return (int) (chars * 256);
    }
}
