package com.insurance.retainedpremium.writer;

import com.insurance.retainedpremium.constant.InsuranceConstants;
import com.insurance.retainedpremium.model.CompanyData;
import com.insurance.retainedpremium.model.QuarterData;
import com.insurance.retainedpremium.reader.ExcelSourceReader;
import com.insurance.retainedpremium.service.DataTransformerService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.*;

class TemplateWriterTest {

    private TemplateWriter templateWriter;
    private ExcelSourceReader excelSourceReader;
    private DataTransformerService dataTransformerService;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        dataTransformerService = new DataTransformerService();
        excelSourceReader = new ExcelSourceReader();
        templateWriter = new TemplateWriter(dataTransformerService);
    }

    @Test
    void writeReport_shouldProduceValidOutput() throws Exception {
        String sourcePath = Path.of("src/test/resources/test-data/29_115(01-03)_自留保費統計表.xlsx")
                .toAbsolutePath().toString();
        Optional<CompanyData> result = excelSourceReader.readSourceFile(sourcePath);
        assertTrue(result.isPresent(), "Should read sample source file");
        CompanyData company29 = result.get();

        Map<String, CompanyData> companies = new HashMap<>();
        companies.put(company29.companyCode(), company29);
        QuarterData q1 = new QuarterData(115, 1, 1, 3, companies);
        Map<Integer, QuarterData> quarterDataMap = Map.of(1, q1);

        Map<String, Double> lastYearData = Map.of("29", 5000000.0);

        String templatePath = Path.of("src/test/resources/test-data/template.xlsx")
                .toAbsolutePath().toString();
        String outputPath = tempDir.resolve("output.xlsx").toString();

        templateWriter.writeReport(templatePath, outputPath, quarterDataMap, lastYearData, 115, 1);

        try (FileInputStream fis = new FileInputStream(outputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            assertEquals("115自留(季)", workbook.getSheetName(0));
            assertEquals("115自留總累", workbook.getSheetName(1));

            XSSFSheet sheet1 = workbook.getSheetAt(0);
            XSSFSheet sheet2 = workbook.getSheetAt(1);

            int[] s1Block = InsuranceConstants.S1_QUARTER_BLOCKS.get(1);
            Integer company29Row = findCompanyRow(sheet1, s1Block[0], s1Block[1], "29");
            assertNotNull(company29Row, "Company 29 should exist in Sheet1 Q1 block");

            Row s1Row = sheet1.getRow(company29Row - 1);
            assertNotNull(s1Row, "Company 29 row should exist in Sheet1");
            assertFalse(s1Row.getZeroHeight(), "Company 29 row should be visible");

            boolean hasData = false;
            for (int col : InsuranceConstants.TEMPLATE_S1_CODE_TO_COL.values()) {
                Cell cell = s1Row.getCell(col - 1);
                if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() != 0.0) {
                    hasData = true;
                    break;
                }
            }
            assertTrue(hasData, "Company 29 should have non-zero data in Sheet1");

            int hiddenCount = 0;
            for (int r = s1Block[0]; r <= s1Block[1]; r++) {
                Row row = sheet1.getRow(r - 1);
                if (row != null && row.getZeroHeight()) {
                    hiddenCount++;
                }
            }
            assertTrue(hiddenCount > 0, "Some rows should be hidden in Sheet1 Q1 (only 1 company has data)");

            int[] s2Block = InsuranceConstants.S2_QUARTER_BLOCKS.get(1);
            Integer company29RowS2 = findCompanyRow(sheet2, s2Block[0], s2Block[1], "29");
            assertNotNull(company29RowS2, "Company 29 should exist in Sheet2 Q1 block");

            Row s2Row = sheet2.getRow(company29RowS2 - 1);
            assertNotNull(s2Row, "Company 29 row should exist in Sheet2");
            assertFalse(s2Row.getZeroHeight(), "Company 29 row should be visible in Sheet2");

            boolean hasS2Formula = false;
            for (int col = 4; col <= 19; col++) {
                Cell cell = s2Row.getCell(col - 1);
                if (cell != null && cell.getCellType() == CellType.FORMULA) {
                    hasS2Formula = true;
                    break;
                }
            }
            assertTrue(hasS2Formula, "Company 29 Sheet2 D-S should preserve cross-sheet formulas");

            Cell uCell = s2Row.getCell(InsuranceConstants.TEMPLATE_S2_LASTYEAR_COL - 1);
            assertNotNull(uCell, "U column cell should exist");
            assertEquals(5000000.0, uCell.getNumericCellValue(), 0.01, "U column should have last year data");

            String title = sheet2.getRow(0).getCell(0).getStringCellValue();
            assertTrue(title.contains("115年度第1季"), "Title should contain year and quarter");
            assertTrue(title.contains("1-3月份"), "Title should contain month range");

            assertEquals("115年度", sheet2.getRow(3).getCell(19).getStringCellValue());
            assertEquals("114年度", sheet2.getRow(3).getCell(20).getStringCellValue());

            int hiddenCountS2 = 0;
            for (int r = s2Block[0]; r <= s2Block[1]; r++) {
                Row row = sheet2.getRow(r - 1);
                if (row != null && row.getZeroHeight()) {
                    hiddenCountS2++;
                }
            }
            assertTrue(hiddenCountS2 > 0, "Some rows should be hidden in Sheet2 Q1");
        }
    }

    @Test
    void writeReport_withNoLastYearData_shouldLeaveUColumnEmpty() throws Exception {
        String sourcePath = Path.of("src/test/resources/test-data/29_115(01-03)_自留保費統計表.xlsx")
                .toAbsolutePath().toString();
        Optional<CompanyData> result = excelSourceReader.readSourceFile(sourcePath);
        assertTrue(result.isPresent());
        CompanyData company29 = result.get();

        Map<String, CompanyData> companies = new HashMap<>();
        companies.put(company29.companyCode(), company29);
        QuarterData q1 = new QuarterData(115, 1, 1, 3, companies);
        Map<Integer, QuarterData> quarterDataMap = Map.of(1, q1);

        Map<String, Double> lastYearData = Map.of();

        String templatePath = Path.of("src/test/resources/test-data/template.xlsx")
                .toAbsolutePath().toString();
        String outputPath = tempDir.resolve("output_no_lastyear.xlsx").toString();

        templateWriter.writeReport(templatePath, outputPath, quarterDataMap, lastYearData, 115, 1);

        try (FileInputStream fis = new FileInputStream(outputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet2 = workbook.getSheetAt(1);
            int[] s2Block = InsuranceConstants.S2_QUARTER_BLOCKS.get(1);
            Integer company29Row = findCompanyRow(sheet2, s2Block[0], s2Block[1], "29");
            assertNotNull(company29Row);

            Row row = sheet2.getRow(company29Row - 1);
            Cell uCell = row.getCell(InsuranceConstants.TEMPLATE_S2_LASTYEAR_COL - 1);
            assertNotNull(uCell, "U column cell should still exist from template");
        }
    }

    @Test
    void writeReport_withDifferentYear_shouldUpdateSheetNames() throws Exception {
        String sourcePath = Path.of("src/test/resources/test-data/29_115(01-03)_自留保費統計表.xlsx")
                .toAbsolutePath().toString();
        Optional<CompanyData> result = excelSourceReader.readSourceFile(sourcePath);
        assertTrue(result.isPresent());
        CompanyData company29 = result.get();

        Map<String, CompanyData> companies = new HashMap<>();
        companies.put(company29.companyCode(), company29);
        QuarterData q1 = new QuarterData(116, 1, 1, 3, companies);
        Map<Integer, QuarterData> quarterDataMap = Map.of(1, q1);
        Map<String, Double> lastYearData = Map.of("29", 1000.0);

        String templatePath = Path.of("src/test/resources/test-data/template.xlsx")
                .toAbsolutePath().toString();
        String outputPath = tempDir.resolve("output_116.xlsx").toString();

        templateWriter.writeReport(templatePath, outputPath, quarterDataMap, lastYearData, 116, 1);

        try (FileInputStream fis = new FileInputStream(outputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            assertEquals("116自留(季)", workbook.getSheetName(0));
            assertEquals("116自留總累", workbook.getSheetName(1));

            XSSFSheet sheet2 = workbook.getSheetAt(1);
            String title = sheet2.getRow(0).getCell(0).getStringCellValue();
            assertTrue(title.startsWith("116年度第1季"), "Title should use year 116");

            assertEquals("116年度", sheet2.getRow(3).getCell(19).getStringCellValue());
            assertEquals("115年度", sheet2.getRow(3).getCell(20).getStringCellValue());
        }
    }

    private Integer findCompanyRow(XSSFSheet sheet, int dataStart, int dataEnd, String companyCode) {
        for (int r = dataStart; r <= dataEnd; r++) {
            Row row = sheet.getRow(r - 1);
            if (row == null) continue;
            Cell cell = row.getCell(0);
            if (cell == null) continue;

            String code;
            if (cell.getCellType() == CellType.NUMERIC) {
                code = String.format("%02d", (int) cell.getNumericCellValue());
            } else if (cell.getCellType() == CellType.STRING) {
                code = cell.getStringCellValue().trim();
            } else {
                continue;
            }
            if (companyCode.equals(code)) return r;
        }
        return null;
    }
}
