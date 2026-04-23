package com.example.retainedpremium.model;

import java.util.Map;

/**
 * 單一公司的自留保費資料
 * retainedPremiums: 險種代號 → 自留保費金額 (e.g., "0100" → 12345678)
 */
public record CompanyData(
    String companyCode,
    String companyName,
    Map<String, Double> retainedPremiums
) {}
