package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.constant.InsuranceConstants;
import com.insurance.retainedpremium.model.CompanyData;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class ExcelSourceReaderTest {

    @Autowired
    private ExcelSourceReader excelSourceReader;

    @Test
    void readSourceFile_shouldParseActualSampleFile() {
        String filePath = Path.of("src/test/resources/test-data/29_115Q1_自留保費統計表.xlsx")
                .toAbsolutePath().toString();

        Optional<CompanyData> result = excelSourceReader.readSourceFile(filePath);

        assertTrue(result.isPresent(), "Should successfully parse the sample file");

        CompanyData data = result.get();

        assertEquals("29", data.companyCode(), "Company code should be 29");

        assertTrue(data.companyName().contains("美國國際"),
                "Company name should contain '美國國際', got: " + data.companyName());

        Map<String, Double> premiums = data.retainedPremiums();
        assertEquals(33, premiums.size(),
                "Should have 33 insurance code entries, got: " + premiums.size());

        for (String code : InsuranceConstants.INSURANCE_CODES) {
            assertTrue(premiums.containsKey(code),
                    "Missing insurance code: " + code);
            assertNotNull(premiums.get(code),
                    "Premium value should not be null for code: " + code);
            assertFalse(Double.isNaN(premiums.get(code)),
                    "Premium value should not be NaN for code: " + code);
        }
    }

    @Test
    void readSourceFile_shouldReturnEmptyForInvalidPath() {
        Optional<CompanyData> result = excelSourceReader.readSourceFile("nonexistent.xlsx");
        assertTrue(result.isEmpty(), "Should return empty for invalid file path");
    }
}
