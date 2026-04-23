package com.example.retainedpremium.service;

import com.example.retainedpremium.model.CompanyData;
import com.example.retainedpremium.model.FileInfo;
import com.example.retainedpremium.model.QuarterData;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class DataTransformerServiceTest {

    private DataTransformerService service;

    @BeforeEach
    void setUp() {
        service = new DataTransformerService();
    }

    @Test
    void groupByQuarter_groupsFilesByQuarter() {
        FileInfo f1 = new FileInfo("A01", 113, 1, 3, 1, "A01_113_Q1.xlsx");
        FileInfo f2 = new FileInfo("A02", 113, 1, 3, 1, "A02_113_Q1.xlsx");
        FileInfo f3 = new FileInfo("B01", 113, 1, 6, 2, "B01_113_Q2.xlsx");

        CompanyData cdA01 = new CompanyData("A01", "Company A01", Map.of());
        CompanyData cdA02 = new CompanyData("A02", "Company A02", Map.of());
        CompanyData cdB01 = new CompanyData("B01", "Company B01", Map.of());

        Map<String, CompanyData> companyDataMap = Map.of(
                "A01", cdA01,
                "A02", cdA02,
                "B01", cdB01
        );

        Map<Integer, QuarterData> result = service.groupByQuarter(List.of(f1, f2, f3), companyDataMap);

        assertEquals(2, result.size());

        QuarterData q1 = result.get(1);
        assertNotNull(q1);
        assertEquals(2, q1.companies().size());
        assertTrue(q1.companies().containsKey("A01"));
        assertTrue(q1.companies().containsKey("A02"));
        assertEquals(113, q1.year());
        assertEquals(1, q1.quarter());

        QuarterData q2 = result.get(2);
        assertNotNull(q2);
        assertEquals(1, q2.companies().size());
        assertTrue(q2.companies().containsKey("B01"));
    }

    @Test
    void groupByQuarter_missingCompanyDataIsSkipped() {
        FileInfo f1 = new FileInfo("A01", 113, 1, 3, 1, "A01_113_Q1.xlsx");
        FileInfo f2 = new FileInfo("MISSING", 113, 1, 3, 1, "MISSING_113_Q1.xlsx");

        CompanyData cdA01 = new CompanyData("A01", "Company A01", Map.of());

        Map<String, CompanyData> companyDataMap = Map.of("A01", cdA01);

        Map<Integer, QuarterData> result = service.groupByQuarter(List.of(f1, f2), companyDataMap);

        assertEquals(1, result.size());
        QuarterData q1 = result.get(1);
        assertNotNull(q1);
        assertEquals(1, q1.companies().size());
        assertTrue(q1.companies().containsKey("A01"));
        assertFalse(q1.companies().containsKey("MISSING"));
    }

    @Test
    void aggregateCategories_sumsCodesIntoCategory() {
        Map<String, Double> premiums = new HashMap<>();
        premiums.put("0100", 100.0);
        premiums.put("0200", 200.0);
        premiums.put("0300", 300.0);
        premiums.put("0400", 400.0);

        CompanyData companyData = new CompanyData("A01", "Company A01", premiums);

        Map<Integer, Double> result = service.aggregateCategories(companyData);

        // 火險 column 4 = 0100 + 0200 + 0300 + 0400 = 1000
        assertEquals(1000.0, result.get(4), 0.001);
    }

    @Test
    void aggregateCategories_allZeroValues() {
        Map<String, Double> premiums = new HashMap<>();
        premiums.put("0100", 0.0);
        premiums.put("0200", 0.0);
        premiums.put("0300", 0.0);
        premiums.put("0400", 0.0);
        premiums.put("0500", 0.0);
        premiums.put("0600", 0.0);
        premiums.put("0700", 0.0);
        premiums.put("0800", 0.0);

        CompanyData companyData = new CompanyData("A01", "Company A01", premiums);

        Map<Integer, Double> result = service.aggregateCategories(companyData);

        for (Double value : result.values()) {
            assertEquals(0.0, value, 0.001);
        }
    }

    @Test
    void aggregateCategories_missingCodesTreatedAsZero() {
        // Empty premiums map - no codes present at all
        Map<String, Double> premiums = new HashMap<>();

        CompanyData companyData = new CompanyData("A01", "Company A01", premiums);

        Map<Integer, Double> result = service.aggregateCategories(companyData);

        assertNotNull(result);
        assertFalse(result.isEmpty());
        for (Double value : result.values()) {
            assertEquals(0.0, value, 0.001);
        }
    }
}
