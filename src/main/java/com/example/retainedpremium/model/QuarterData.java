package com.example.retainedpremium.model;

import java.util.Map;

/**
 * 單一季度的所有公司資料
 * companies: 公司代號 → CompanyData
 */
public record QuarterData(
    int year,
    int quarter,
    int startMonth,
    int endMonth,
    Map<String, CompanyData> companies
) {}
