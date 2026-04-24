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
 * 對照範例檔:
 *   背景: 淺青色 #CCFFFF (表頭列 + 左側固定欄)
 *   數值格式: #,##0_);[Red](#,##0)
 *   百分比: 0.00%
 */
public class ExcelStyleHelper {

    private static final byte[] CYAN_BG = {(byte) 0xCC, (byte) 0xFF, (byte) 0xFF};
    private static final byte[] DARK_BLUE = {0x00, 0x00, (byte) 0x80};

    // === Row Heights (in points) ===
    public static final float S1_ROW0_HEIGHT = 17.25f;
    public static final float S1_ROW1_HEIGHT = 34.5f;
    public static final float S1_DATA_ROW_HEIGHT = 22.5f;
    public static final float S2_ROW0_HEIGHT = 31.5f;
    public static final float S2_ROW1_HEIGHT = 18.0f;
    public static final float S2_ROW2_HEIGHT = 21.75f;
    public static final float S2_ROW3_HEIGHT = 27.0f;
    public static final float S2_ROW4_HEIGHT = 21.75f;
    public static final float S2_DATA_ROW_HEIGHT = 45.0f;

    // === Sheet1 樣式 ===
    private CellStyle s1CodeHeaderStyle;     // Row0: 險種代號標籤 (全真標準楷書 10pt, 青色底)
    private CellStyle s1CodeValueStyle;      // Row0: 險種代號值 (Book Antiqua 12pt, center/bottom)
    private CellStyle s1CodeValueBoldStyle;  // Row0: B1空欄 (Book Antiqua 12pt bold, center/bottom)
    private CellStyle s1NameHeaderStyle;     // Row1: D+險種名稱 (全真標準楷書 11pt, 青色底, 框線, center)
    private CellStyle s1CompanyNameHeaderStyle; // Row1: C欄公司別/險種 (全真標準楷書 11pt, 青色底, 框線)
    private CellStyle s1LabelHeaderStyle;    // Row1: A欄標題 (全真標準楷書 10pt, 青色底, 框線, left)
    private CellStyle s1MonthHeaderStyle;    // Row1: B欄"月份" (全真標準楷書 10pt bold, 深藍, 青色底, 框線, center)
    private CellStyle s1CompanyCodeStyle;    // 公司代號 (Book Antiqua 10pt, 青色底, 框線, center)
    private CellStyle s1MonthStyle;          // 月份 (Book Antiqua 10pt bold, 深藍, 青色底, 框線, center, text)
    private CellStyle s1CompanyNameStyle;    // 公司名稱 (全真標準楷書 12pt, 青色底, 框線)
    private CellStyle s1DataStyle;           // 數值 (Book Antiqua 12pt, #,##0, 框線)
    private CellStyle s1SubtotalCodeStyle;   // 小計代號欄 (Book Antiqua 10pt, 青色底, 框線)
    private CellStyle s1SubtotalMonthStyle;  // 小計月份欄 (Book Antiqua 10pt bold, 深藍, 青色底, 框線, text)
    private CellStyle s1SubtotalLabelStyle;  // 小計標籤 (全真標準楷書 12pt, 青色底, 框線)
    private CellStyle s1SubtotalDataStyle;   // 小計數值 (Book Antiqua 12pt, #,##0, 框線)

    // === Sheet2 樣式 ===
    private CellStyle s2TitleStyle;          // 標題 (DFKai-SB 24pt 粗體, center)
    private CellStyle s2UnitStyle;           // 單位 (DFKai-SB 10pt)
    private CellStyle s2GroupHeaderStyle;    // Row2 群組標題 (PMingLiu 15pt 粗體, 青色底, 框線)
    private CellStyle s2HeaderCodeStyle;     // Row3 A欄代號 (PMingLiu 12pt 粗體, 青色底, 框線)
    private CellStyle s2HeaderCompanyStyle;  // Row3 B/C欄 (PMingLiu 14pt 粗體, 青色底, 框線)
    private CellStyle s2MainHeaderStyle;     // Row3 D+主標題 (PMingLiu 15pt 粗體, 青色底, 框線)
    private CellStyle s2SubHeaderStyle;      // Row4 子標題 (PMingLiu 15pt 粗體, 青色底, 框線)
    private CellStyle s2CompanyCodeStyle;    // A欄代號 (Book Antiqua 10pt, 青色底, 框線, center)
    private CellStyle s2MonthStyle;          // B欄月份 (Book Antiqua 10pt bold, 青色底, 框線, center)
    private CellStyle s2CompanyNameStyle;    // C欄公司名 (全真標準楷書 14pt, 青色底, 框線)
    private CellStyle s2DataStyle;           // 數值 (Book Antiqua 14pt, #,##0, 框線, vcenter)
    private CellStyle s2PercentStyle;        // 百分比 (Book Antiqua 14pt, 0.00%, 框線)
    private CellStyle s2SubtotalCodeStyle;   // 小計A欄 (Book Antiqua 10pt, 青色底, 框線)
    private CellStyle s2SubtotalMonthStyle;  // 小計B欄 (Book Antiqua 10pt bold, 青色底, 框線)
    private CellStyle s2SubtotalLabelStyle;  // 小計C欄 (全真標準楷書 14pt, 青色底, 框線)
    private CellStyle s2SubtotalDataStyle;   // 小計數值 (Book Antiqua 14pt, #,##0, 框線)
    private CellStyle s2SubtotalPctStyle;    // 小計百分比 (Book Antiqua 14pt, 0.00%, 框線)

    // === Sheet3 (歸屬) 樣式 ===
    private CellStyle guishuHeaderStyle;     // 標頭 (PMingLiu 16pt)
    private CellStyle guishuCategoryStyle;   // 分類 (PMingLiu 14pt 粗體)
    private CellStyle guishuDataStyle;       // 資料 (PMingLiu 14pt)
    private CellStyle guishuCodeStyle;       // 代號 (Book Antiqua 14pt)

    public ExcelStyleHelper(Workbook workbook) {
        initStyles((XSSFWorkbook) workbook);
    }

    private void initStyles(XSSFWorkbook wb) {
        XSSFColor cyanColor = new XSSFColor(CYAN_BG, null);
        XSSFColor darkBlueColor = new XSSFColor(DARK_BLUE, null);
        DataFormat df = wb.createDataFormat();
        short numberFmt = df.getFormat("#,##0_);[Red]\\(#,##0\\)");
        short percentFmt = df.getFormat("0.00%");
        short textFmt = df.getFormat("@");

        // === 字體定義 ===
        XSSFFont bookAntiqua12 = createFont(wb, "Book Antiqua", 12, false);
        XSSFFont bookAntiqua12bold = createFont(wb, "Book Antiqua", 12, true);
        XSSFFont bookAntiqua10 = createFont(wb, "Book Antiqua", 10, false);
        XSSFFont bookAntiqua10bold = createFont(wb, "Book Antiqua", 10, true);
        bookAntiqua10bold.setColor(darkBlueColor);
        XSSFFont bookAntiqua14 = createFont(wb, "Book Antiqua", 14, false);
        XSSFFont quanzhen11 = createFont(wb, "全真標準楷書", 11, false);
        XSSFFont quanzhen10 = createFont(wb, "全真標準楷書", 10, false);
        XSSFFont quanzhen10bold = createFont(wb, "全真標準楷書", 10, true);
        quanzhen10bold.setColor(darkBlueColor);
        XSSFFont quanzhen12 = createFont(wb, "全真標準楷書", 12, false);
        XSSFFont quanzhen14 = createFont(wb, "全真標準楷書", 14, false);
        XSSFFont dfkaisb24bold = createFont(wb, "DFKai-SB", 24, true);
        XSSFFont dfkaisb10 = createFont(wb, "DFKai-SB", 10, false);
        XSSFFont pmingliu15bold = createFont(wb, "PMingLiu", 15, true);
        XSSFFont pmingliu16 = createFont(wb, "PMingLiu", 16, false);
        XSSFFont pmingliu14bold = createFont(wb, "PMingLiu", 14, true);
        XSSFFont pmingliu14 = createFont(wb, "PMingLiu", 14, false);
        XSSFFont pmingliu12bold = createFont(wb, "PMingLiu", 12, true);

        // === Sheet1 樣式 ===

        // Row0: "險種代號" label (全真標準楷書 10pt, left/center)
        s1CodeHeaderStyle = wb.createCellStyle();
        s1CodeHeaderStyle.setFont(quanzhen10);
        s1CodeHeaderStyle.setAlignment(HorizontalAlignment.LEFT);
        s1CodeHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Row0: 險種代號值 (Book Antiqua 12pt, center/bottom)
        s1CodeValueStyle = wb.createCellStyle();
        s1CodeValueStyle.setFont(bookAntiqua12);
        s1CodeValueStyle.setAlignment(HorizontalAlignment.CENTER);
        s1CodeValueStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

        // Row0: B1空欄 (Book Antiqua 12pt bold, center/bottom)
        s1CodeValueBoldStyle = wb.createCellStyle();
        s1CodeValueBoldStyle.setFont(bookAntiqua12bold);
        s1CodeValueBoldStyle.setAlignment(HorizontalAlignment.CENTER);
        s1CodeValueBoldStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

        // Row1: 險種名稱 D-AK (全真標準楷書 11pt, center/center, cyan, borders)
        s1NameHeaderStyle = wb.createCellStyle();
        s1NameHeaderStyle.setFont(quanzhen11);
        s1NameHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s1NameHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s1NameHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) s1NameHeaderStyle, cyanColor);
        applyThinBorders(s1NameHeaderStyle);

        // Row1: C欄 "公司別/險種" (全真標準楷書 11pt, vcenter, cyan, borders, NO halign)
        s1CompanyNameHeaderStyle = wb.createCellStyle();
        s1CompanyNameHeaderStyle.setFont(quanzhen11);
        s1CompanyNameHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s1CompanyNameHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) s1CompanyNameHeaderStyle, cyanColor);
        applyThinBorders(s1CompanyNameHeaderStyle);

        // Row1: A欄 "代號"(全真標準楷書 10pt, left/center, cyan, borders)
        s1LabelHeaderStyle = wb.createCellStyle();
        s1LabelHeaderStyle.setFont(quanzhen10);
        s1LabelHeaderStyle.setAlignment(HorizontalAlignment.LEFT);
        s1LabelHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s1LabelHeaderStyle, cyanColor);
        applyThinBorders(s1LabelHeaderStyle);

        // Row1: B欄 "月份" (全真標準楷書 10pt bold, dark blue, center/center, cyan, borders, text)
        s1MonthHeaderStyle = wb.createCellStyle();
        s1MonthHeaderStyle.setFont(quanzhen10bold);
        s1MonthHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s1MonthHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s1MonthHeaderStyle.setDataFormat(textFmt);
        applyCyanBg((XSSFCellStyle) s1MonthHeaderStyle, cyanColor);
        applyThinBorders(s1MonthHeaderStyle);

        // Data A: 公司代號 (Book Antiqua 10pt, center/bottom, cyan, borders)
        s1CompanyCodeStyle = wb.createCellStyle();
        s1CompanyCodeStyle.setFont(bookAntiqua10);
        s1CompanyCodeStyle.setAlignment(HorizontalAlignment.CENTER);
        s1CompanyCodeStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        applyCyanBg((XSSFCellStyle) s1CompanyCodeStyle, cyanColor);
        applyThinBorders(s1CompanyCodeStyle);

        // Data B: 月份 (Book Antiqua 10pt bold, dark blue, center/bottom, cyan, borders, text)
        s1MonthStyle = wb.createCellStyle();
        s1MonthStyle.setFont(bookAntiqua10bold);
        s1MonthStyle.setAlignment(HorizontalAlignment.CENTER);
        s1MonthStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        s1MonthStyle.setDataFormat(textFmt);
        applyCyanBg((XSSFCellStyle) s1MonthStyle, cyanColor);
        applyThinBorders(s1MonthStyle);

        // Data C: 公司名稱 (全真標準楷書 12pt, vcenter, cyan, borders)
        s1CompanyNameStyle = wb.createCellStyle();
        s1CompanyNameStyle.setFont(quanzhen12);
        s1CompanyNameStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s1CompanyNameStyle, cyanColor);
        applyThinBorders(s1CompanyNameStyle);

        // Data D+: 數值 (Book Antiqua 12pt, #,##0, borders, bottom)
        s1DataStyle = wb.createCellStyle();
        s1DataStyle.setFont(bookAntiqua12);
        s1DataStyle.setDataFormat(numberFmt);
        s1DataStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        applyThinBorders(s1DataStyle);

        // Subtotal A: empty code cell (Book Antiqua 10pt, cyan, borders)
        s1SubtotalCodeStyle = wb.createCellStyle();
        s1SubtotalCodeStyle.setFont(bookAntiqua10);
        s1SubtotalCodeStyle.setAlignment(HorizontalAlignment.CENTER);
        s1SubtotalCodeStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        applyCyanBg((XSSFCellStyle) s1SubtotalCodeStyle, cyanColor);
        applyThinBorders(s1SubtotalCodeStyle);

        // Subtotal B: month label (same as data month style)
        s1SubtotalMonthStyle = wb.createCellStyle();
        s1SubtotalMonthStyle.setFont(bookAntiqua10bold);
        s1SubtotalMonthStyle.setAlignment(HorizontalAlignment.CENTER);
        s1SubtotalMonthStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        s1SubtotalMonthStyle.setDataFormat(textFmt);
        applyCyanBg((XSSFCellStyle) s1SubtotalMonthStyle, cyanColor);
        applyThinBorders(s1SubtotalMonthStyle);

        // Subtotal C: "小計" label (全真標準楷書 12pt, cyan, borders)
        s1SubtotalLabelStyle = wb.createCellStyle();
        s1SubtotalLabelStyle.setFont(quanzhen12);
        s1SubtotalLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s1SubtotalLabelStyle, cyanColor);
        applyThinBorders(s1SubtotalLabelStyle);

        // Subtotal data (Book Antiqua 12pt, #,##0, borders)
        s1SubtotalDataStyle = wb.createCellStyle();
        s1SubtotalDataStyle.setFont(bookAntiqua12);
        s1SubtotalDataStyle.setDataFormat(numberFmt);
        applyThinBorders(s1SubtotalDataStyle);

        // === Sheet2 樣式 ===
        s2TitleStyle = wb.createCellStyle();
        s2TitleStyle.setFont(dfkaisb24bold);
        s2TitleStyle.setAlignment(HorizontalAlignment.CENTER);
        s2TitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 單位 (DFKai-SB 10pt, right, vcenter)
        s2UnitStyle = wb.createCellStyle();
        s2UnitStyle.setFont(dfkaisb10);
        s2UnitStyle.setAlignment(HorizontalAlignment.RIGHT);
        s2UnitStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Row2: 群組 (PMingLiu 15pt bold, cyan, borders, center)
        s2GroupHeaderStyle = wb.createCellStyle();
        s2GroupHeaderStyle.setFont(pmingliu15bold);
        s2GroupHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s2GroupHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2GroupHeaderStyle, cyanColor);
        applyThinBorders(s2GroupHeaderStyle);

        // Row3: A欄代號 (PMingLiu 12pt bold, cyan, borders, center)
        s2HeaderCodeStyle = wb.createCellStyle();
        s2HeaderCodeStyle.setFont(pmingliu12bold);
        s2HeaderCodeStyle.setAlignment(HorizontalAlignment.CENTER);
        s2HeaderCodeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2HeaderCodeStyle, cyanColor);
        applyThinBorders(s2HeaderCodeStyle);

        // Row3: B/C欄 (PMingLiu 14pt bold, cyan, borders, center)
        s2HeaderCompanyStyle = wb.createCellStyle();
        s2HeaderCompanyStyle.setFont(pmingliu14bold);
        s2HeaderCompanyStyle.setAlignment(HorizontalAlignment.CENTER);
        s2HeaderCompanyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s2HeaderCompanyStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) s2HeaderCompanyStyle, cyanColor);
        applyThinBorders(s2HeaderCompanyStyle);

        // Row3: 主標題(PMingLiu 15pt bold, cyan, borders, center, wrap)
        s2MainHeaderStyle = wb.createCellStyle();
        s2MainHeaderStyle.setFont(pmingliu15bold);
        s2MainHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s2MainHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        s2MainHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) s2MainHeaderStyle, cyanColor);
        applyThinBorders(s2MainHeaderStyle);

        // Row4: 子標題 (PMingLiu 15pt bold, cyan, borders, center)
        s2SubHeaderStyle = wb.createCellStyle();
        s2SubHeaderStyle.setFont(pmingliu15bold);
        s2SubHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        s2SubHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2SubHeaderStyle, cyanColor);
        applyThinBorders(s2SubHeaderStyle);

        // Data A: 代號 (Book Antiqua 10pt, center/center, cyan, borders)
        s2CompanyCodeStyle = wb.createCellStyle();
        s2CompanyCodeStyle.setFont(bookAntiqua10);
        s2CompanyCodeStyle.setAlignment(HorizontalAlignment.CENTER);
        s2CompanyCodeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2CompanyCodeStyle, cyanColor);
        applyThinBorders(s2CompanyCodeStyle);

        // Data B: 月份 (Book Antiqua 10pt bold, center/center, cyan, borders)
        s2MonthStyle = wb.createCellStyle();
        s2MonthStyle.setFont(bookAntiqua10bold);
        s2MonthStyle.setAlignment(HorizontalAlignment.CENTER);
        s2MonthStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2MonthStyle, cyanColor);
        applyThinBorders(s2MonthStyle);

        // Data C: 公司名 (全真標準楷書 14pt, vcenter, cyan, borders)
        s2CompanyNameStyle = wb.createCellStyle();
        s2CompanyNameStyle.setFont(quanzhen14);
        s2CompanyNameStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2CompanyNameStyle, cyanColor);
        applyThinBorders(s2CompanyNameStyle);

        // Data D-U: 數值 (Book Antiqua 14pt, #,##0, vcenter, borders)
        s2DataStyle = wb.createCellStyle();
        s2DataStyle.setFont(bookAntiqua14);
        s2DataStyle.setDataFormat(numberFmt);
        s2DataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s2DataStyle);

        // Data V: 百分比 (Book Antiqua 14pt, 0.00%, borders, center)
        s2PercentStyle = wb.createCellStyle();
        s2PercentStyle.setFont(bookAntiqua14);
        s2PercentStyle.setDataFormat(percentFmt);
        s2PercentStyle.setAlignment(HorizontalAlignment.CENTER);
        s2PercentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s2PercentStyle);

        // Subtotal A (Book Antiqua 10pt, cyan, borders)
        s2SubtotalCodeStyle = wb.createCellStyle();
        s2SubtotalCodeStyle.setFont(bookAntiqua10);
        s2SubtotalCodeStyle.setAlignment(HorizontalAlignment.CENTER);
        s2SubtotalCodeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2SubtotalCodeStyle, cyanColor);
        applyThinBorders(s2SubtotalCodeStyle);

        // Subtotal B (Book Antiqua 10pt bold, cyan, borders)
        s2SubtotalMonthStyle = wb.createCellStyle();
        s2SubtotalMonthStyle.setFont(bookAntiqua10bold);
        s2SubtotalMonthStyle.setAlignment(HorizontalAlignment.CENTER);
        s2SubtotalMonthStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2SubtotalMonthStyle, cyanColor);
        applyThinBorders(s2SubtotalMonthStyle);

        // Subtotal C: "小計" (全真標準楷書 14pt, cyan, borders)
        s2SubtotalLabelStyle = wb.createCellStyle();
        s2SubtotalLabelStyle.setFont(quanzhen14);
        s2SubtotalLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) s2SubtotalLabelStyle, cyanColor);
        applyThinBorders(s2SubtotalLabelStyle);

        // Subtotal data (Book Antiqua 14pt, #,##0, borders)
        s2SubtotalDataStyle = wb.createCellStyle();
        s2SubtotalDataStyle.setFont(bookAntiqua14);
        s2SubtotalDataStyle.setDataFormat(numberFmt);
        s2SubtotalDataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s2SubtotalDataStyle);

        // Subtotal percent (Book Antiqua 14pt, 0.00%, borders, center)
        s2SubtotalPctStyle = wb.createCellStyle();
        s2SubtotalPctStyle.setFont(bookAntiqua14);
        s2SubtotalPctStyle.setDataFormat(percentFmt);
        s2SubtotalPctStyle.setAlignment(HorizontalAlignment.CENTER);
        s2SubtotalPctStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(s2SubtotalPctStyle);

        // === Sheet3 (歸屬) 樣式 ===
        guishuHeaderStyle = wb.createCellStyle();
        guishuHeaderStyle.setFont(pmingliu16);
        guishuHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        guishuHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        guishuCategoryStyle = wb.createCellStyle();
        guishuCategoryStyle.setFont(pmingliu14bold);
        guishuCategoryStyle.setAlignment(HorizontalAlignment.CENTER);
        guishuCategoryStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

        guishuDataStyle = wb.createCellStyle();
        guishuDataStyle.setFont(pmingliu14);
        guishuDataStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

        guishuCodeStyle = wb.createCellStyle();
        guishuCodeStyle.setFont(bookAntiqua14);
        guishuCodeStyle.setAlignment(HorizontalAlignment.CENTER);
        guishuCodeStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
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
    public CellStyle getS1CodeValueStyle() { return s1CodeValueStyle; }
    public CellStyle getS1CodeValueBoldStyle() { return s1CodeValueBoldStyle; }
    public CellStyle getS1NameHeaderStyle() { return s1NameHeaderStyle; }
    public CellStyle getS1CompanyNameHeaderStyle() { return s1CompanyNameHeaderStyle; }
    public CellStyle getS1LabelHeaderStyle() { return s1LabelHeaderStyle; }
    public CellStyle getS1MonthHeaderStyle() { return s1MonthHeaderStyle; }
    public CellStyle getS1CompanyCodeStyle() { return s1CompanyCodeStyle; }
    public CellStyle getS1MonthStyle() { return s1MonthStyle; }
    public CellStyle getS1CompanyNameStyle() { return s1CompanyNameStyle; }
    public CellStyle getS1DataStyle() { return s1DataStyle; }
    public CellStyle getS1SubtotalCodeStyle() { return s1SubtotalCodeStyle; }
    public CellStyle getS1SubtotalMonthStyle() { return s1SubtotalMonthStyle; }
    public CellStyle getS1SubtotalLabelStyle() { return s1SubtotalLabelStyle; }
    public CellStyle getS1SubtotalDataStyle() { return s1SubtotalDataStyle; }

    // === Sheet2 getters ===
    public CellStyle getS2TitleStyle() { return s2TitleStyle; }
    public CellStyle getS2UnitStyle() { return s2UnitStyle; }
    public CellStyle getS2GroupHeaderStyle() { return s2GroupHeaderStyle; }
    public CellStyle getS2HeaderCodeStyle() { return s2HeaderCodeStyle; }
    public CellStyle getS2HeaderCompanyStyle() { return s2HeaderCompanyStyle; }
    public CellStyle getS2MainHeaderStyle() { return s2MainHeaderStyle; }
    public CellStyle getS2SubHeaderStyle() { return s2SubHeaderStyle; }
    public CellStyle getS2CompanyCodeStyle() { return s2CompanyCodeStyle; }
    public CellStyle getS2MonthStyle() { return s2MonthStyle; }
    public CellStyle getS2CompanyNameStyle() { return s2CompanyNameStyle; }
    public CellStyle getS2DataStyle() { return s2DataStyle; }
    public CellStyle getS2PercentStyle() { return s2PercentStyle; }
    public CellStyle getS2SubtotalCodeStyle() { return s2SubtotalCodeStyle; }
    public CellStyle getS2SubtotalMonthStyle() { return s2SubtotalMonthStyle; }
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
