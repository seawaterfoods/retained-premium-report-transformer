package com.example.retainedpremium.constant;

import java.util.*;

public final class InsuranceConstants {
    private InsuranceConstants() {}

    // ==================== 33 Insurance Type Codes ====================
    // Source file rows 7-39, column A
    public static final List<String> INSURANCE_CODES = List.of(
        "0100", "0200", "0300", "0400", "0500", "0600", "0700", "0800", "0900",
        "1000", "1100", "1200", "1300", "1400", "1500", "1600",
        "1700", "1800", "1900", "2000", "2100", "2200", "2300",
        "2400", "2500", "2600", "2700", "2800", "2900",
        "3000", "3100", "3200", "9900"
    );

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

    // ==================== Template Sheet1: Insurance Code → Column ====================
    // Sheet1 columns: D=4, E=5, ..., AJ=36, AK=37
    // Maps insurance code to 1-based column index in template Sheet1
    public static final Map<String, Integer> TEMPLATE_S1_CODE_TO_COL;
    static {
        Map<String, Integer> map = new LinkedHashMap<>();
        map.put("0100", 4);  map.put("0200", 5);  map.put("0300", 6);  map.put("0400", 7);
        map.put("0500", 8);  map.put("0600", 9);  map.put("0700", 10); map.put("0800", 11);
        map.put("0900", 12); map.put("1000", 13); map.put("1100", 14); map.put("1200", 15);
        map.put("1300", 16); map.put("1400", 17); map.put("1500", 18); map.put("1600", 19);
        map.put("1700", 20); map.put("1800", 21); map.put("1900", 22); map.put("2000", 23);
        map.put("2100", 24); map.put("2200", 25); map.put("2300", 26); map.put("2400", 27);
        map.put("2500", 28); map.put("2600", 29); map.put("2700", 30); map.put("2800", 31);
        map.put("2900", 32); map.put("3000", 33); map.put("3100", 34); map.put("3200", 35);
        map.put("9900", 36);
        TEMPLATE_S1_CODE_TO_COL = Collections.unmodifiableMap(map);
    }
    public static final int TEMPLATE_S1_TOTAL_COL = 37; // AK (formula, do NOT write)

    // ==================== Template Sheet1: Quarter Blocks (1-based rows) ====================
    // Each quarter has 19 data rows + 1 subtotal row
    // dataStart = first company row, dataEnd = last company row, subtotal = SUBTOTAL formula row
    public static final int S1_Q1_DATA_START = 3;   public static final int S1_Q1_DATA_END = 21;  public static final int S1_Q1_SUBTOTAL = 22;
    public static final int S1_Q2_DATA_START = 23;  public static final int S1_Q2_DATA_END = 41;  public static final int S1_Q2_SUBTOTAL = 42;
    public static final int S1_Q3_DATA_START = 43;  public static final int S1_Q3_DATA_END = 61;  public static final int S1_Q3_SUBTOTAL = 62;
    public static final int S1_Q4_DATA_START = 63;  public static final int S1_Q4_DATA_END = 81;  public static final int S1_Q4_SUBTOTAL = 82;

    // Structured access: quarter (1-4) → {dataStart, dataEnd, subtotal}
    public static final Map<Integer, int[]> S1_QUARTER_BLOCKS;
    static {
        Map<Integer, int[]> map = new HashMap<>();
        map.put(1, new int[]{S1_Q1_DATA_START, S1_Q1_DATA_END, S1_Q1_SUBTOTAL});
        map.put(2, new int[]{S1_Q2_DATA_START, S1_Q2_DATA_END, S1_Q2_SUBTOTAL});
        map.put(3, new int[]{S1_Q3_DATA_START, S1_Q3_DATA_END, S1_Q3_SUBTOTAL});
        map.put(4, new int[]{S1_Q4_DATA_START, S1_Q4_DATA_END, S1_Q4_SUBTOTAL});
        S1_QUARTER_BLOCKS = Collections.unmodifiableMap(map);
    }

    // ==================== Template Sheet2: Quarter Blocks (1-based rows) ====================
    public static final int S2_Q1_DATA_START = 6;   public static final int S2_Q1_DATA_END = 24;  public static final int S2_Q1_SUBTOTAL = 25;
    public static final int S2_Q2_DATA_START = 26;  public static final int S2_Q2_DATA_END = 44;  public static final int S2_Q2_SUBTOTAL = 45;
    public static final int S2_Q3_DATA_START = 46;  public static final int S2_Q3_DATA_END = 64;  public static final int S2_Q3_SUBTOTAL = 65;
    public static final int S2_Q4_DATA_START = 66;  public static final int S2_Q4_DATA_END = 84;  public static final int S2_Q4_SUBTOTAL = 85;
    public static final int S2_GRAND_TOTAL_ROW = 86;  // Grand total
    public static final int S2_COUNT_ROW = 87;         // Company count row

    public static final Map<Integer, int[]> S2_QUARTER_BLOCKS;
    static {
        Map<Integer, int[]> map = new HashMap<>();
        map.put(1, new int[]{S2_Q1_DATA_START, S2_Q1_DATA_END, S2_Q1_SUBTOTAL});
        map.put(2, new int[]{S2_Q2_DATA_START, S2_Q2_DATA_END, S2_Q2_SUBTOTAL});
        map.put(3, new int[]{S2_Q3_DATA_START, S2_Q3_DATA_END, S2_Q3_SUBTOTAL});
        map.put(4, new int[]{S2_Q4_DATA_START, S2_Q4_DATA_END, S2_Q4_SUBTOTAL});
        S2_QUARTER_BLOCKS = Collections.unmodifiableMap(map);
    }

    // ==================== Sheet2: Category Mapping ====================
    // Category name → (1-based column index, list of insurance codes that sum into this category)
    // D=4 火險, E=5 水險, F=6 航空, G=7 車體損失, H=8 任意車責, I=9 強制車責,
    // J=10 強制機車, K=11 二輪電動, L=12 責任險, M=13 工程險, N=14 信用保證,
    // O=15 其他財產, P=16 傷害險, Q=17 天災險, R=18 健康險, S=19 國外分進

    // Use a record-like structure: CategoryDef(columnIndex, codes)
    public static final Map<String, CategoryDef> CATEGORY_MAPPING;
    static {
        Map<String, CategoryDef> map = new LinkedHashMap<>();
        map.put("火險",       new CategoryDef(4,  List.of("0100", "0200", "0300", "0400")));
        map.put("水險",       new CategoryDef(5,  List.of("0500", "0600", "0700", "0800")));
        map.put("航空",       new CategoryDef(6,  List.of("0900")));
        map.put("車體損失險", new CategoryDef(7,  List.of("1000", "1100")));
        map.put("任意責任險", new CategoryDef(8,  List.of("1200", "1300")));
        map.put("強制車責",   new CategoryDef(9,  List.of("1400", "1500")));
        map.put("強制機車",   new CategoryDef(10, List.of("1600")));
        map.put("二輪電動",   new CategoryDef(11, List.of("3200")));
        map.put("責任險",     new CategoryDef(12, List.of("1700", "1800")));
        map.put("工程險",     new CategoryDef(13, List.of("1900")));
        map.put("信用保證",   new CategoryDef(14, List.of("2100", "2200")));
        map.put("其他財產",   new CategoryDef(15, List.of("2000", "2300", "2600", "2700")));
        map.put("傷害險",     new CategoryDef(16, List.of("2400")));
        map.put("天災險",     new CategoryDef(17, List.of("2500", "2800", "2900")));
        map.put("健康險",     new CategoryDef(18, List.of("3000", "3100")));
        map.put("國外分進",   new CategoryDef(19, List.of("9900")));
        CATEGORY_MAPPING = Collections.unmodifiableMap(map);
    }

    public static final int TEMPLATE_S2_YEAR_TOTAL_COL = 20;    // T (formula, do NOT write)
    public static final int TEMPLATE_S2_LASTYEAR_COL = 21;      // U (write last year data)
    public static final int TEMPLATE_S2_GROWTH_COL = 22;        // V (formula, do NOT write)

    // ==================== Quarter Labels ====================
    // quarter → label string used in column B of both sheets
    public static final Map<Integer, String> QUARTER_LABEL_MAP = Map.of(
        1, "1-1Q(1-3)",
        2, "1-2Q(1-6)",
        3, "1-3Q(1-9)",
        4, "1-4Q(1-12)"
    );

    // quarter → (startMonth, endMonth)
    public static final Map<Integer, int[]> QUARTER_MONTH_MAP = Map.of(
        1, new int[]{1, 3},
        2, new int[]{1, 6},
        3, new int[]{1, 9},
        4, new int[]{1, 12}
    );

    // ==================== CategoryDef record ====================
    public record CategoryDef(int columnIndex, List<String> insuranceCodes) {}
}
