package com.insurance.retainedpremium.config;

import com.insurance.retainedpremium.config.AppConfig.*;
import jakarta.annotation.PostConstruct;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 險種歸屬對照服務 — 從 config/insurance-mapping.yml 載入設定。
 * <p>
 * 提供險種代號、名稱、分類對照等資料結構，取代原本硬編碼的 InsuranceConstants。
 */
@Component
public class InsuranceMappingService {

    private static final Logger log = LoggerFactory.getLogger(InsuranceMappingService.class);

    private final AppConfig config;

    // 險種代號清單 (順序 = Sheet1 欄位順序)
    private List<String> insuranceCodes;
    private Map<String, String> codeToShortName;
    private Map<String, String> codeToFullName;

    // Sheet1 欄位對應 (code → 1-based column)
    private Map<String, Integer> s1CodeToCol;
    private int s1ColDataEnd;
    private int s1ColTotal;

    // Sheet2 分類對應
    private Map<String, CategoryDef> categoryMapping;
    private int s2ColDataEnd;
    private int s2ColYearTotal;
    private int s2ColLastYear;
    private int s2ColGrowth;

    // Sheet2 群組標頭 (Row3 merged)
    private Map<String, int[]> s2CategoryGroups;
    // Sheet2 子標頭 (Row5)
    private Map<Integer, String> s2SubHeaders;
    // Sheet2 強制責任險合併範圍 (Row4)
    private List<int[]> s2HeaderGroupMerges;

    // Sheet3 歸屬表
    private List<GuishuEntry> guishuTable;

    public InsuranceMappingService(AppConfig config) {
        this.config = config;
    }

    @PostConstruct
    public void init() {
        InsuranceConfig ins = config.getInsurance();

        // 險種代號
        insuranceCodes = new ArrayList<>();
        codeToShortName = new LinkedHashMap<>();
        codeToFullName = new LinkedHashMap<>();
        for (CodeConfig cc : ins.getCodes()) {
            insuranceCodes.add(cc.getCode());
            codeToShortName.put(cc.getCode(), cc.getShortName());
            codeToFullName.put(cc.getCode(), cc.getFullName());
        }
        insuranceCodes = Collections.unmodifiableList(insuranceCodes);
        codeToShortName = Collections.unmodifiableMap(codeToShortName);
        codeToFullName = Collections.unmodifiableMap(codeToFullName);

        // Sheet1 column mapping (D 欄起始 = column 4, 1-based)
        s1CodeToCol = new LinkedHashMap<>();
        int col = 4; // S1_COL_DATA_START
        for (String code : insuranceCodes) {
            s1CodeToCol.put(code, col++);
        }
        s1CodeToCol = Collections.unmodifiableMap(s1CodeToCol);
        s1ColDataEnd = col - 1;
        s1ColTotal = col;

        // 分類、群組標頭、子標頭、歸屬表
        categoryMapping = new LinkedHashMap<>();
        s2CategoryGroups = new LinkedHashMap<>();
        s2SubHeaders = new LinkedHashMap<>();
        s2HeaderGroupMerges = new ArrayList<>();
        guishuTable = new ArrayList<>();

        int s2Col = 4; // S2_COL_DATA_START
        Map<String, int[]> groupTracker = new LinkedHashMap<>();

        // 中文數字序號
        String[] chineseNumbers = {"一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
                "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"};
        int guishuIdx = 0;

        for (MajorCategoryConfig mc : ins.getCategories()) {
            String majorName = mc.getName();
            String majorNumber = mc.getNumber();

            // 記錄此大類在 Sheet2 的起始欄
            int majorStartCol = s2Col;

            for (SubCategoryConfig sc : mc.getSubCategories()) {
                String subName = sc.getName();
                categoryMapping.put(subName, new CategoryDef(s2Col, List.copyOf(sc.getCodes())));

                // 子標頭 (Row5)
                if (sc.getHeaderLabel() != null) {
                    s2SubHeaders.put(s2Col, sc.getHeaderLabel());
                }
                if (sc.getSubHeader() != null) {
                    s2SubHeaders.put(s2Col, sc.getSubHeader());
                }

                // 群組標頭合併追蹤
                if (sc.getHeaderGroup() != null) {
                    int currentS2Col = s2Col;
                    groupTracker.computeIfAbsent(sc.getHeaderGroup(), k -> new int[]{currentS2Col, currentS2Col});
                    groupTracker.get(sc.getHeaderGroup())[1] = s2Col;
                }

                // 歸屬表
                boolean isFirst = true;
                for (String code : sc.getCodes()) {
                    if (isFirst) {
                        guishuTable.add(new GuishuEntry(
                                guishuIdx < chineseNumbers.length ? chineseNumbers[guishuIdx] : String.valueOf(guishuIdx + 1),
                                subName, code, codeToFullName.getOrDefault(code, "")));
                        guishuIdx++;
                        isFirst = false;
                    } else {
                        guishuTable.add(new GuishuEntry(null, null, code,
                                codeToFullName.getOrDefault(code, "")));
                    }
                }

                s2Col++;
            }

            int majorEndCol = s2Col - 1;

            // 大類群組 (多於1個子分類時 Row3 合併)
            if (majorEndCol > majorStartCol) {
                s2CategoryGroups.put(majorName, new int[]{majorStartCol, majorEndCol});
            }
        }

        // 處理 headerGroup 合併 (如強制責任險)
        for (var entry : groupTracker.entrySet()) {
            s2HeaderGroupMerges.add(entry.getValue());
        }

        s2ColDataEnd = s2Col - 1;
        s2ColYearTotal = s2Col;
        s2ColLastYear = s2Col + 1;
        s2ColGrowth = s2Col + 2;

        // 比較欄群組
        s2CategoryGroups.put("比較", new int[]{s2ColGrowth, s2ColGrowth});

        categoryMapping = Collections.unmodifiableMap(categoryMapping);
        s2CategoryGroups = Collections.unmodifiableMap(s2CategoryGroups);
        s2SubHeaders = Collections.unmodifiableMap(s2SubHeaders);
        s2HeaderGroupMerges = Collections.unmodifiableList(s2HeaderGroupMerges);
        guishuTable = Collections.unmodifiableList(guishuTable);

        log.info("險種歸屬載入完成: {} 代號, {} 子分類, {} 歸屬項目",
                insuranceCodes.size(), categoryMapping.size(), guishuTable.size());
    }

    // --- Getters ---

    public List<String> getInsuranceCodes() { return insuranceCodes; }
    public Map<String, String> getCodeToShortName() { return codeToShortName; }
    public Map<String, String> getCodeToFullName() { return codeToFullName; }
    public Map<String, Integer> getS1CodeToCol() { return s1CodeToCol; }
    public int getS1ColDataEnd() { return s1ColDataEnd; }
    public int getS1ColTotal() { return s1ColTotal; }
    public Map<String, CategoryDef> getCategoryMapping() { return categoryMapping; }
    public int getS2ColDataEnd() { return s2ColDataEnd; }
    public int getS2ColYearTotal() { return s2ColYearTotal; }
    public int getS2ColLastYear() { return s2ColLastYear; }
    public int getS2ColGrowth() { return s2ColGrowth; }
    public Map<String, int[]> getS2CategoryGroups() { return s2CategoryGroups; }
    public Map<Integer, String> getS2SubHeaders() { return s2SubHeaders; }
    public List<int[]> getS2HeaderGroupMerges() { return s2HeaderGroupMerges; }
    public List<GuishuEntry> getGuishuTable() { return guishuTable; }

    /** 分類定義 */
    public record CategoryDef(int columnIndex, List<String> insuranceCodes) {}

    /** 歸屬表項目 */
    public record GuishuEntry(String category, String group, String code, String name) {}
}
