package com.insurance.retainedpremium.constant;

import java.util.*;

public final class InsuranceConstants {
    private InsuranceConstants() {}

    // ==================== 33 Insurance Type Codes ====================
    public static final List<String> INSURANCE_CODES = List.of(
        "0100", "0200", "0300", "0400", "0500", "0600", "0700", "0800", "0900",
        "1000", "1100", "1200", "1300", "1400", "1500", "1600",
        "1700", "1800", "1900", "2000", "2100", "2200", "2300",
        "2400", "2500", "2600", "2700", "2800", "2900",
        "3000", "3100", "3200", "9900"
    );

    // ==================== Insurance Code → Name (Sheet1 Row2 headers) ====================
    public static final Map<String, String> INSURANCE_CODE_NAMES;
    static {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("0100", "一年期住宅火險");
        map.put("0200", "長期住宅火險");
        map.put("0300", "一年期商業火險");
        map.put("0400", "長期商業火險");
        map.put("0500", "內陸運輸險");
        map.put("0600", "貨物運輸險");
        map.put("0700", "船體險");
        map.put("0800", "漁船險");
        map.put("0900", "航空險");
        map.put("1000", "自用車損險");
        map.put("1100", "商業車損險");
        map.put("1200", "自用車責險");
        map.put("1300", "商業車責險");
        map.put("1400", "強制自用車責險");
        map.put("1500", "強制商業車責險");
        map.put("1600", "強制機車責險");
        map.put("1700", "一般責任險");
        map.put("1800", "專業責任險");
        map.put("1900", "工程險");
        map.put("2000", "核能險");
        map.put("2100", "保證險");
        map.put("2200", "信用險");
        map.put("2300", "其他財產保險");
        map.put("2400", "傷害險");
        map.put("2500", "商業地震險");
        map.put("2600", "個人綜險");
        map.put("2700", "商業綜險");
        map.put("2800", "颱風洪水險");
        map.put("2900", "政策地震險");
        map.put("3000", "一年期健康險");
        map.put("3100", "長年期健康險");
        map.put("3200", "強制電動二輪");
        map.put("9900", "國外分進");
        INSURANCE_CODE_NAMES = Collections.unmodifiableMap(map);
    }

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

    // ==================== Sheet1 Column Layout (1-based) ====================
    public static final int S1_COL_CODE = 1;       // A: 公司代號
    public static final int S1_COL_MONTH = 2;      // B: 月份
    public static final int S1_COL_COMPANY = 3;    // C: 公司別/險種
    public static final int S1_COL_DATA_START = 4; // D: first insurance code (0100)
    public static final int S1_COL_DATA_END = 36;  // AJ: last insurance code (9900)
    public static final int S1_COL_TOTAL = 37;     // AK: 合計

    // Insurance Code → Sheet1 Column (1-based)
    public static final Map<String, Integer> S1_CODE_TO_COL;
    static {
        Map<String, Integer> map = new LinkedHashMap<>();
        int col = S1_COL_DATA_START;
        for (String code : INSURANCE_CODES) {
            map.put(code, col++);
        }
        S1_CODE_TO_COL = Collections.unmodifiableMap(map);
    }

    // ==================== Sheet1 Header Rows (0-based) ====================
    public static final int S1_HEADER_ROW_CODES = 0;  // Row 1: insurance codes
    public static final int S1_HEADER_ROW_NAMES = 1;  // Row 2: insurance type names
    public static final int S1_DATA_START_ROW = 2;     // Row 3: first data row (0-based)

    // ==================== Sheet2 Column Layout (1-based) ====================
    public static final int S2_COL_CODE = 1;          // A
    public static final int S2_COL_MONTH = 2;         // B
    public static final int S2_COL_COMPANY = 3;       // C
    public static final int S2_COL_DATA_START = 4;    // D: first category (火險)
    public static final int S2_COL_DATA_END = 19;     // S: last category (國外分進)
    public static final int S2_COL_YEAR_TOTAL = 20;   // T: 年度合計
    public static final int S2_COL_LASTYEAR = 21;     // U: 去年同期
    public static final int S2_COL_GROWTH = 22;       // V: 成長率

    // ==================== Sheet2 Header Rows (0-based) ====================
    public static final int S2_HEADER_ROW_TITLE = 0;  // Row 1: title
    public static final int S2_HEADER_ROW_UNIT = 1;   // Row 2: unit
    public static final int S2_HEADER_ROW_GROUP = 2;   // Row 3: category groups (汽車險, 意外險)
    public static final int S2_HEADER_ROW_MAIN = 3;    // Row 4: main headers
    public static final int S2_HEADER_ROW_SUB = 4;     // Row 5: sub headers
    public static final int S2_DATA_START_ROW = 5;     // Row 6: first data row (0-based)

    // ==================== Sheet2: Category Mapping ====================
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

    // Sheet2 category group headers (Row 3 merged regions)
    public static final Map<String, int[]> S2_CATEGORY_GROUPS = Map.of(
        "汽車險", new int[]{7, 11},   // G-K
        "意外險", new int[]{12, 15}   // L-O
    );

    // Sheet2 sub-headers for 強制責任險 breakdown (Row 5)
    public static final Map<Integer, String> S2_SUB_HEADERS = Map.of(
        9, "汽車",
        10, "機車",
        11, "二輪電動",
        15, "責任保險"
    );

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

    // ==================== Sheet3 (歸屬) Category Attribution ====================
    public record GuishuEntry(String category, String group, String code, String name) {}

    public static final List<GuishuEntry> GUISHU_TABLE = List.of(
        new GuishuEntry("一", "火險", "0100", "一年期住宅火災保險"),
        new GuishuEntry(null,  null,   "0200", "長期住宅火災保險"),
        new GuishuEntry(null,  null,   "0300", "一年期商業火災保險"),
        new GuishuEntry(null,  null,   "0400", "長期商業火災保險"),
        new GuishuEntry("二", "水險", "0500", "內陸運輸保險"),
        new GuishuEntry(null,  null,   "0600", "貨物運輸保險"),
        new GuishuEntry(null,  null,   "0700", "船體保險"),
        new GuishuEntry(null,  null,   "0800", "漁船保險"),
        new GuishuEntry("三", "航空", "0900", "航空保險"),
        new GuishuEntry("四", "車體損失險", "1000", "自用車損險"),
        new GuishuEntry(null,  null,   "1100", "商業車損險"),
        new GuishuEntry("五", "任意責任險", "1200", "自用車責險"),
        new GuishuEntry(null,  null,   "1300", "商業車責險"),
        new GuishuEntry("六", "強制車責",   "1400", "強制自用車責險"),
        new GuishuEntry(null,  null,   "1500", "強制商業車責險"),
        new GuishuEntry("七", "強制機車",   "1600", "強制機車責險"),
        new GuishuEntry("八", "二輪電動",   "3200", "強制電動二輪"),
        new GuishuEntry("九", "責任險",     "1700", "一般責任險"),
        new GuishuEntry(null,  null,   "1800", "專業責任險"),
        new GuishuEntry("十", "工程險",     "1900", "工程保險"),
        new GuishuEntry("十一", "信用保證", "2100", "保證保險"),
        new GuishuEntry(null,  null,   "2200", "信用保險"),
        new GuishuEntry("十二", "其他財產", "2000", "核能保險"),
        new GuishuEntry(null,  null,   "2300", "其他財產保險"),
        new GuishuEntry(null,  null,   "2600", "個人綜合保險"),
        new GuishuEntry(null,  null,   "2700", "商業綜合保險"),
        new GuishuEntry("十三", "傷害險",   "2400", "傷害保險"),
        new GuishuEntry("十四", "天災險",   "2500", "商業地震保險"),
        new GuishuEntry(null,  null,   "2800", "颱風洪水保險"),
        new GuishuEntry(null,  null,   "2900", "政策地震保險"),
        new GuishuEntry("十五", "健康險",   "3000", "一年期健康險"),
        new GuishuEntry(null,  null,   "3100", "長年期健康險"),
        new GuishuEntry("十六", "國外分進", "9900", "國外分進")
    );

    // ==================== CategoryDef record ====================
    public record CategoryDef(int columnIndex, List<String> insuranceCodes) {}
}
