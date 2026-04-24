package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.model.CompanyData;
import com.insurance.retainedpremium.model.FileInfo;
import com.insurance.retainedpremium.model.ValidationResult;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

class FileValidatorTest {

    private FileValidator validator;

    @BeforeEach
    void setUp() {
        validator = new FileValidator();
    }

    private Map<String, Double> buildPremiums(int count, double value) {
        return IntStream.rangeClosed(1, count)
                .boxed()
                .collect(Collectors.toMap(
                        i -> String.format("%04d", i * 100),
                        i -> value
                ));
    }

    @Test
    void validateFileInfos_validFiles_passed() {
        List<FileInfo> fileInfos = List.of(
                new FileInfo("A01", 2024, 1, 3, 1, "A01_2024_Q1.xlsx"),
                new FileInfo("A02", 2024, 4, 6, 2, "A02_2024_Q2.xlsx")
        );

        ValidationResult result = validator.validateFileInfos(fileInfos);

        assertTrue(result.isPassed());
        assertTrue(result.getErrors().isEmpty());
    }

    @Test
    void validateFileInfos_mixedYears_error() {
        List<FileInfo> fileInfos = List.of(
                new FileInfo("A01", 2024, 1, 3, 1, "A01_2024_Q1.xlsx"),
                new FileInfo("A02", 2023, 4, 6, 2, "A02_2023_Q2.xlsx")
        );

        ValidationResult result = validator.validateFileInfos(fileInfos);

        assertFalse(result.isPassed());
        assertFalse(result.getErrors().isEmpty());
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("Inconsistent year")));
    }

    @Test
    void validateFileInfos_duplicateCompanySameQuarter_error() {
        List<FileInfo> fileInfos = List.of(
                new FileInfo("A01", 2024, 1, 3, 1, "A01_2024_Q1.xlsx"),
                new FileInfo("A01", 2024, 1, 3, 1, "A01_2024_Q1_v2.xlsx")
        );

        ValidationResult result = validator.validateFileInfos(fileInfos);

        assertFalse(result.isPassed());
        assertFalse(result.getErrors().isEmpty());
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("Duplicate company")));
    }

    @Test
    void validateFileInfos_emptyList_error() {
        ValidationResult result = validator.validateFileInfos(Collections.emptyList());

        assertFalse(result.isPassed());
        assertFalse(result.getErrors().isEmpty());
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("No files")));
    }

    @Test
    void validateCompanyData_validData_passed() {
        Map<String, Double> premiums = buildPremiums(33, 1000.0);
        CompanyData data = new CompanyData("A01", "Company A", premiums);

        ValidationResult result = validator.validateCompanyData(data, "A01.xlsx");

        assertTrue(result.isPassed());
        assertTrue(result.getErrors().isEmpty());
        assertTrue(result.getWarnings().isEmpty());
    }

    @Test
    void validateCompanyData_nanValue_error() {
        Map<String, Double> premiums = new HashMap<>(buildPremiums(33, 1000.0));
        premiums.put("0100", Double.NaN);
        CompanyData data = new CompanyData("A01", "Company A", premiums);

        ValidationResult result = validator.validateCompanyData(data, "A01.xlsx");

        assertFalse(result.isPassed());
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("Invalid numeric value")));
    }

    @Test
    void validateCompanyData_negativeValue_warningOnly() {
        Map<String, Double> premiums = new HashMap<>(buildPremiums(33, 1000.0));
        premiums.put("0100", -500.0);
        CompanyData data = new CompanyData("A01", "Company A", premiums);

        ValidationResult result = validator.validateCompanyData(data, "A01.xlsx");

        assertTrue(result.isPassed());
        assertTrue(result.getErrors().isEmpty());
        assertFalse(result.getWarnings().isEmpty());
        assertTrue(result.getWarnings().stream().anyMatch(w -> w.contains("Negative")));
    }

    @Test
    void validateCompanyData_emptyCompanyCode_error() {
        Map<String, Double> premiums = buildPremiums(33, 1000.0);
        CompanyData data = new CompanyData("", "Company A", premiums);

        ValidationResult result = validator.validateCompanyData(data, "test.xlsx");

        assertFalse(result.isPassed());
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("Company code is empty")));
    }
}
