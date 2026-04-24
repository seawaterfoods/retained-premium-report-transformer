package com.insurance.retainedpremium.constant;

import java.util.Map;

/**
 * 靜態常數 — 來源檔案佈局、Sheet 標頭列、季度標籤。
 * <p>
 * 險種代號、分類對照、歸屬表等動態設定請見 {@code InsuranceMappingService}。
 */
public final class InsuranceConstants {
    private InsuranceConstants() {}

    // ==================== Source File Layout ====================
    public static final int SOURCE_DATA_START_ROW = 7;   // 1-based
    public static final int SOURCE_DATA_END_ROW = 39;    // 1-based
    public static final int SOURCE_COMPANY_CODE_ROW = 4; // B4
    public static final int SOURCE_COMPANY_CODE_COL = 2; // B
    public static final int SOURCE_COMPANY_NAME_ROW = 4; // C4
    public static final int SOURCE_COMPANY_NAME_COL = 3; // C
    public static final int SOURCE_RETAINED_COL = 6;     // F (retained premium)
    public static final int SOURCE_SIGNED_COL = 3;       // C (signed premium)
    public static final int SOURCE_REINSURANCE_IN_COL = 4;  // D
    public static final int SOURCE_REINSURANCE_OUT_COL = 5; // E
    public static final int SOURCE_CODE_COL = 1;         // A (insurance code)

    // ==================== Sheet1 Fixed Columns (1-based) ====================
    public static final int S1_COL_CODE = 1;       // A: 公司代號
    public static final int S1_COL_MONTH = 2;      // B: 月份
    public static final int S1_COL_COMPANY = 3;    // C: 公司別/險種
    public static final int S1_COL_DATA_START = 4; // D: first insurance code

    // ==================== Sheet1 Header Rows (0-based) ====================
    public static final int S1_HEADER_ROW_CODES = 0;  // Row 1: insurance codes
    public static final int S1_HEADER_ROW_NAMES = 1;  // Row 2: insurance type names
    public static final int S1_DATA_START_ROW = 2;     // Row 3: first data row (0-based)

    // ==================== Sheet2 Fixed Columns (1-based) ====================
    public static final int S2_COL_CODE = 1;          // A
    public static final int S2_COL_MONTH = 2;         // B
    public static final int S2_COL_COMPANY = 3;       // C
    public static final int S2_COL_DATA_START = 4;    // D: first category

    // ==================== Sheet2 Header Rows (0-based) ====================
    public static final int S2_HEADER_ROW_TITLE = 0;  // Row 1: title
    public static final int S2_HEADER_ROW_UNIT = 1;   // Row 2: unit
    public static final int S2_HEADER_ROW_GROUP = 2;   // Row 3: category groups
    public static final int S2_HEADER_ROW_MAIN = 3;    // Row 4: main headers
    public static final int S2_HEADER_ROW_SUB = 4;     // Row 5: sub headers
    public static final int S2_DATA_START_ROW = 5;     // Row 6: first data row (0-based)

    // ==================== Quarter Labels ====================
    public static final Map<Integer, String> QUARTER_LABEL_MAP = Map.of(
        1, "1-1Q(1-3)",
        2, "1-2Q(1-6)",
        3, "1-3Q(1-9)",
        4, "1-4Q(1-12)"
    );

    public static final Map<Integer, int[]> QUARTER_MONTH_MAP = Map.of(
        1, new int[]{1, 3},
        2, new int[]{1, 6},
        3, new int[]{1, 9},
        4, new int[]{1, 12}
    );
}
