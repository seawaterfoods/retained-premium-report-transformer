package com.insurance.retainedpremium.model;

/**
 * 來源檔名解析結果
 */
public record FileInfo(
    String companyCode,
    int year,
    int startMonth,
    int endMonth,
    int quarter,
    String filename
) {}
