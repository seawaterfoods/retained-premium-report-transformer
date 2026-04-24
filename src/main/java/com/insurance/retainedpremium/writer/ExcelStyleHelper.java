package com.insurance.retainedpremium.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel 格式工具 — 集中管理所有儲存格樣式
 *
 * 對照範例模板:
 *   背景: 淺青色 #CCFFFF (表頭列)
 *   數值格式: #,##0_);[Red](#,##0)
 *   百分比: 0.00%
 */
public class ExcelStyleHelper {

    private static final byte[] CYAN_BG = {(byte) 0xCC, (byte) 0xFF, (byte) 0xFF};

    // === Sheet1 樣式 ===
    private CellStyle s1CodeHeaderStyle;     // Row1: 險種代號 (Book Antiqua 12pt)
    private CellStyle s1NameHeaderStyle;     // Row2: 險種名稱 (全真標準楷書 11pt, 青色底, 框線)
    private CellStyle s1LabelHeaderStyle;    // Row2: A-C 欄標題 (全真標準楷書 10pt, 青色底, 框線)
    private CellStyle s1CompanyCodeStyle;    // 公司代號 (Book Antiqua 10pt, 框線)
    private CellStyle s1MonthStyle;          // 月份 (Book Antiqua 10pt, 框線)
    private CellStyle s1CompanyNameStyle;    // 公司名稱 (全真標準楷書 12pt, 框線)
    private CellStyle s1DataStyle;           // 數值 (Book Antiqua 12pt, #,##0, 框線)
    private CellStyle s1SubtotalLabelStyle;  // 小計標籤 (全真標準楷書 12pt, 框線)
    private CellStyle s1SubtotalDataStyle;   // 小計數值 (Book Antiqua 12pt, #,##0, 框線)

    // === Sheet2 樣式 ===
    private CellStyle s2TitleStyle;          // 標題 (DFKai-SB 24pt 粗體)
    private CellStyle s2UnitStyle;           // 單位 (10pt)
    private CellStyle s2GroupHeaderStyle;    // Row3 群組標題 (PMingLiu 15pt 粗體, 青色底)
    private CellStyle s2MainHeaderStyle;     // Row4 主標題 (PMingLiu 15pt 粗體, 青色底, 框線)
    private CellStyle s2SubHeaderStyle;      // Row5 子標題 (PMingLiu 12pt, 青色底, 框線)
    private CellStyle s2LabelStyle;          // 代號/月份/公司 (PMingLiu 12pt, 框線)
    private CellStyle s2DataStyle;           // 數值 (Book Antiqua 14pt, #,##0, 框線)
    private CellStyle s2PercentStyle;        // 百分比 (Book Antiqua 14pt, 0.00%, 框線)
    private CellStyle s2SubtotalLabelStyle;  // 小計標籤 (PMingLiu 12pt, 框線)
    private CellStyle s2SubtotalDataStyle;   // 小計數值 (Book Antiqua 14pt, #,##0, 框線)
    private CellStyle s2SubtotalPctStyle;    // 小計百分比 (Book Antiqua 14pt, 0.00%, 框線)

    // === Sheet3 (歸屬) 樣式 ===
    private CellStyle guishuHeaderStyle;     // 標頭 (新細明體 16pt)
    private CellStyle guishuCategoryStyle;   // 分類 (新細明體 14pt 粗體)
    private CellStyle guishuDataStyle;       // 資料 (新細明體 14pt)
    private CellStyle guishuCodeStyle;       // 代號 (Book Antiqua 14pt)

    public ExcelStyleHelper(Workbook workbook) {
        initStyles((XSSFWorkbook) workbook);
    }

    private void initStyles(XSSFWorkbook wb) {
        XSSFColor cyanColor = new XSSFColor(CYAN_BG, null);
        DataFormat df = wb.createDataFormat();
        short numberFmt = df.getFormat("#,##0_);[Red]\\(#,##0\\)");
        short percentFmt = df.getFormat("0.00%");

        // === 字體定義 ===
        XSSFFont bookAntiqua12 = createFont(wb, "Book Antiqua", 12, false);
        XSSFFont bookAntiqua10 = createFont(wb, "Book Antiqua", 10, false);
        XSSFFont bookAntiqua14 = createFont(wb, "Book Antiqua", 14, false);
        XSSFFont quanzhen11 = createFont(wb, "全真標準楷書", 11, false);
        XSSFFont quanzhen10 = createFont(wb, "全真標準楷書", 10, false);
        XSSFFont quanzhen12 = createFont(wb, "全真標準楷書", 12, false);
        XSSFFont dfkaisb24bold = createFont(wb, "DFKai-SB", 24, true);
        XSSFFont pmingliu15bold = createFont(wb, "PMingLiu", 15, true);
        XSSFFont pmingliu12 = createFont(wb, "PMingLiu", 12, false);
        XSSFFont pmingliu10 = createFont(wb, "PMingLiu", 10, false);
        XSSFFont sinmingti16 = createFont(wb, "新細明體", 16, false);
        XSSFFont sinmingti14bold = createFont(wb, "新細明體", 14, true);
        XSSFFont sinmingti14 = createFont(wb, "新細明體", 14, false);

        // === Sheet1 樣式 ===
        s1CodeHeaderStyle = wb.createCellStyle();
        s1CodeHeaderStyle.setFont(bookAntiqua12);
        s1CodeHeaderStyle.setAlignment(HorizontalAlignment.LEFT);
        s1CodeHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        s1NameHeaderStyle = wb.createCellStyle();
        s1NameHeaderStyle.setFont(quanzhen11);
        s1NameHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s1NameHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s1NameHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) s1NameHeaderStyle, cyanColor);
        applyThinBorders(s1NameHeaderStyle);

        s1LabelHeaderStyle = wb.createCellStyle();
        s1LabelHeaderStyle.setFont(quanzhen10);
        s1LabelHeaderStyle.setAlignment(HorizontalAlignment.LEFT);
        s1LabelHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s1LabelHeaderStyle, cyanColor);
        applyThinBorders(s1LabelHeaderStyle);

        s1CompanyCodeStyle = wb.createCellStyle();
        s1CompanyCodeStyle.setFont(bookAntiqua10);
        s1CompanyCodeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s1CompanyCodeStyle);

        s1MonthStyle = wb.createCellStyle();
        s1MonthStyle.setFont(bookAntiqua10);
        s1MonthStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s1MonthStyle);

        s1CompanyNameStyle = wb.createCellStyle();
        s1CompanyNameStyle.setFont(quanzhen12);
        s1CompanyNameStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s1CompanyNameStyle);

        s1DataStyle = wb.createCellStyle();
        s1DataStyle.setFont(bookAntiqua12);
        s1DataStyle.setDataFormat(numberFmt);
        s1DataStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(s1DataStyle);

        s1SubtotalLabelStyle = wb.createCellStyle();
        s1SubtotalLabelStyle.setFont(quanzhen12);
        s1SubtotalLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s1SubtotalLabelStyle);

        s1SubtotalDataStyle = wb.createCellStyle();
        s1SubtotalDataStyle.setFont(bookAntiqua12);
        s1SubtotalDataStyle.setDataFormat(numberFmt);
        s1SubtotalDataStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(s1SubtotalDataStyle);

        // === Sheet2 樣式 ===
        s2TitleStyle = wb.createCellStyle();
        s2TitleStyle.setFont(dfkaisb24bold);
        s2TitleStyle.setAlignment(HorizontalAlignment.CENTER);
        s2TitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        s2UnitStyle = wb.createCellStyle();
        s2UnitStyle.setFont(pmingliu10);
        s2UnitStyle.setAlignment(HorizontalAlignment.RIGHT);

        s2GroupHeaderStyle = wb.createCellStyle();
        s2GroupHeaderStyle.setFont(pmingliu15bold);
        s2GroupHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s2GroupHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2GroupHeaderStyle, cyanColor);
        applyThinBorders(s2GroupHeaderStyle);

        s2MainHeaderStyle = wb.createCellStyle();
        s2MainHeaderStyle.setFont(pmingliu15bold);
        s2MainHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s2MainHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s2MainHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) s2MainHeaderStyle, cyanColor);
        applyThinBorders(s2MainHeaderStyle);

        s2SubHeaderStyle = wb.createCellStyle();
        s2SubHeaderStyle.setFont(pmingliu12);
        s2SubHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s2SubHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2SubHeaderStyle, cyanColor);
        applyThinBorders(s2SubHeaderStyle);

        s2LabelStyle = wb.createCellStyle();
        s2LabelStyle.setFont(pmingliu12);
        s2LabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s2LabelStyle);

        s2DataStyle = wb.createCellStyle();
        s2DataStyle.setFont(bookAntiqua14);
        s2DataStyle.setDataFormat(numberFmt);
        s2DataStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(s2DataStyle);

        s2PercentStyle = wb.createCellStyle();
        s2PercentStyle.setFont(bookAntiqua14);
        s2PercentStyle.setDataFormat(percentFmt);
        s2PercentStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(s2PercentStyle);

        s2SubtotalLabelStyle = wb.createCellStyle();
        s2SubtotalLabelStyle.setFont(pmingliu12);
        s2SubtotalLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s2SubtotalLabelStyle);

        s2SubtotalDataStyle = wb.createCellStyle();
        s2SubtotalDataStyle.setFont(bookAntiqua14);
        s2SubtotalDataStyle.setDataFormat(numberFmt);
        s2SubtotalDataStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(s2SubtotalDataStyle);

        s2SubtotalPctStyle = wb.createCellStyle();
        s2SubtotalPctStyle.setFont(bookAntiqua14);
        s2SubtotalPctStyle.setDataFormat(percentFmt);
        s2SubtotalPctStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(s2SubtotalPctStyle);

        // === Sheet3 (歸屬) 樣式 ===
        guishuHeaderStyle = wb.createCellStyle();
        guishuHeaderStyle.setFont(sinmingti16);
        guishuHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        guishuCategoryStyle = wb.createCellStyle();
        guishuCategoryStyle.setFont(sinmingti14bold);
        guishuCategoryStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        guishuDataStyle = wb.createCellStyle();
        guishuDataStyle.setFont(sinmingti14);
        guishuDataStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        guishuCodeStyle = wb.createCellStyle();
        guishuCodeStyle.setFont(bookAntiqua14);
        guishuCodeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    private XSSFFont createFont(XSSFWorkbook wb, String name, int size, boolean bold) {
        XSSFFont font = wb.createFont();
        font.setFontName(name);
        font.setFontHeightInPoints((short) size);
        font.setBold(bold);
        return font;
    }

    private void applyCyanBg(XSSFCellStyle style, XSSFColor color) {
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void applyThinBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    // === Sheet1 getters ===
    public CellStyle getS1CodeHeaderStyle() { return s1CodeHeaderStyle; }
    public CellStyle getS1NameHeaderStyle() { return s1NameHeaderStyle; }
    public CellStyle getS1LabelHeaderStyle() { return s1LabelHeaderStyle; }
    public CellStyle getS1CompanyCodeStyle() { return s1CompanyCodeStyle; }
    public CellStyle getS1MonthStyle() { return s1MonthStyle; }
    public CellStyle getS1CompanyNameStyle() { return s1CompanyNameStyle; }
    public CellStyle getS1DataStyle() { return s1DataStyle; }
    public CellStyle getS1SubtotalLabelStyle() { return s1SubtotalLabelStyle; }
    public CellStyle getS1SubtotalDataStyle() { return s1SubtotalDataStyle; }

    // === Sheet2 getters ===
    public CellStyle getS2TitleStyle() { return s2TitleStyle; }
    public CellStyle getS2UnitStyle() { return s2UnitStyle; }
    public CellStyle getS2GroupHeaderStyle() { return s2GroupHeaderStyle; }
    public CellStyle getS2MainHeaderStyle() { return s2MainHeaderStyle; }
    public CellStyle getS2SubHeaderStyle() { return s2SubHeaderStyle; }
    public CellStyle getS2LabelStyle() { return s2LabelStyle; }
    public CellStyle getS2DataStyle() { return s2DataStyle; }
    public CellStyle getS2PercentStyle() { return s2PercentStyle; }
    public CellStyle getS2SubtotalLabelStyle() { return s2SubtotalLabelStyle; }
    public CellStyle getS2SubtotalDataStyle() { return s2SubtotalDataStyle; }
    public CellStyle getS2SubtotalPctStyle() { return s2SubtotalPctStyle; }

    // === Sheet3 getters ===
    public CellStyle getGuishuHeaderStyle() { return guishuHeaderStyle; }
    public CellStyle getGuishuCategoryStyle() { return guishuCategoryStyle; }
    public CellStyle getGuishuDataStyle() { return guishuDataStyle; }
    public CellStyle getGuishuCodeStyle() { return guishuCodeStyle; }

    // === 工具方法 ===

    public static Cell createCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        if (value != null) cell.setCellValue(value);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    public static Cell createCell(Row row, int col, double value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    public static Cell createFormulaCell(Row row, int col, String formula, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellFormula(formula);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    /** 0-based 欄位索引 → Excel 欄位字母 */
    public static String colLetter(int colIndex) {
        StringBuilder sb = new StringBuilder();
        int idx = colIndex;
        while (idx >= 0) {
            sb.insert(0, (char) ('A' + idx % 26));
            idx = idx / 26 - 1;
        }
        return sb.toString();
    }

    /** 產生儲存格參考 (e.g., "D6"), col/row 皆 0-based */
    public static String cellRef(int col, int row) {
        return colLetter(col) + (row + 1);
    }

    public static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        if (firstRow == lastRow && firstCol == lastCol) return;
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }
}
