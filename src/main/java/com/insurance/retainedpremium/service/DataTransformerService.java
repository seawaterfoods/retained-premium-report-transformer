package com.insurance.retainedpremium.service;

import com.insurance.retainedpremium.constant.InsuranceConstants;
import com.insurance.retainedpremium.constant.InsuranceConstants.CategoryDef;
import com.insurance.retainedpremium.model.CompanyData;
import com.insurance.retainedpremium.model.FileInfo;
import com.insurance.retainedpremium.model.QuarterData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class DataTransformerService {

    private static final Logger log = LoggerFactory.getLogger(DataTransformerService.class);

    public Map<Integer, QuarterData> groupByQuarter(List<FileInfo> fileInfos, Map<String, CompanyData> companyDataMap) {
        Map<Integer, Map<String, CompanyData>> quarterCompanies = new HashMap<>();
        Map<Integer, FileInfo> quarterRepresentative = new HashMap<>();

        for (FileInfo fileInfo : fileInfos) {
            CompanyData companyData = companyDataMap.get(fileInfo.companyCode());
            if (companyData == null) {
                log.warn("No company data found for companyCode={}, skipping file: {}", fileInfo.companyCode(), fileInfo.filename());
                continue;
            }
            int quarter = fileInfo.quarter();
            quarterCompanies.computeIfAbsent(quarter, k -> new LinkedHashMap<>())
                    .put(fileInfo.companyCode(), companyData);
            quarterRepresentative.putIfAbsent(quarter, fileInfo);
        }

        Map<Integer, QuarterData> result = new HashMap<>();
        for (var entry : quarterCompanies.entrySet()) {
            int quarter = entry.getKey();
            FileInfo rep = quarterRepresentative.get(quarter);
            result.put(quarter, new QuarterData(rep.year(), quarter, rep.startMonth(), rep.endMonth(), entry.getValue()));
        }
        return result;
    }

    public Map<Integer, Double> aggregateCategories(CompanyData companyData) {
        Map<Integer, Double> result = new HashMap<>();

        for (Map.Entry<String, CategoryDef> entry : InsuranceConstants.CATEGORY_MAPPING.entrySet()) {
            CategoryDef def = entry.getValue();
            double sum = 0.0;
            for (String code : def.insuranceCodes()) {
                Double val = companyData.retainedPremiums().get(code);
                if (val != null) {
                    sum += val;
                }
            }
            result.put(def.columnIndex(), sum);
        }

        return result;
    }
}
