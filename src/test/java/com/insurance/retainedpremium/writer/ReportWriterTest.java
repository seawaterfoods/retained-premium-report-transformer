package com.insurance.retainedpremium.writer;

import com.insurance.retainedpremium.constant.InsuranceConstants;
import com.insurance.retainedpremium.model.CompanyData;
import com.insurance.retainedpremium.model.QuarterData;
import com.insurance.retainedpremium.reader.ExcelSourceReader;
import com.insurance.retainedpremium.service.DataTransformerService;
import org.apache.poi.ss.usermodel.*;
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

class ReportWriterTest {

    private ReportWriter reportWriter;
    private ExcelSourceReader excelSourceReader;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        DataTransformerService dataTransformerService = new DataTransformerService();
        excelSourceReader = new ExcelSourceReader();
        reportWriter = new ReportWriter(dataTransformerService);
    }

    @Test
    void writeReport_shouldProduceValidOutput() throws Exception {
        // Arrange
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

        String outputPath = tempDir.resolve("output.xlsx").toString();

        // Act
        reportWriter.writeReport(outputPath, quarterDataMap, lastYearData, 115, 1);

        // Assert
        try (FileInputStream fis = new FileInputStream(outputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            assertEquals(3, workbook.getNumberOfSheets(), "Should have 3 sheets");
            assertEquals("115自留(季)", workbook.getSheetName(0));
            assertEquals("115自留總累", workbook.getSheetName(1));
            assertEquals("歸屬", workbook.getSheetName(2));

            // === Sheet1 驗證 ===
            XSSFSheet sheet1 = workbook.getSheetAt(0);

            // Header row 0: insurance codes
            Row headerRow0 = sheet1.getRow(0);
            assertNotNull(headerRow0);
            assertEquals("險種代號", headerRow0.getCell(2).getStringCellValue());
            assertEquals("0100", headerRow0.getCell(3).getStringCellValue());

            // Header row 1: names
            Row headerRow1 = sheet1.getRow(1);
            assertNotNull(headerRow1);
            assertEquals("代號", headerRow1.getCell(0).getStringCellValue());
            assertEquals("月份", headerRow1.getCell(1).getStringCellValue());
            assertEquals("公司別/險種", headerRow1.getCell(2).getStringCellValue());
            assertEquals("合計", headerRow1.getCell(36).getStringCellValue());

            // Data row (row 2, 0-based) should have company 29
            Row dataRow = sheet1.getRow(2);
            assertNotNull(dataRow);
            assertEquals("29", dataRow.getCell(0).getStringCellValue());
            assertEquals("1-1Q(1-3)", dataRow.getCell(1).getStringCellValue());
            assertFalse(dataRow.getZeroHeight(), "Company 29 should be visible");

            // Check that data values exist (non-zero for at least some codes)
            boolean hasData = false;
            for (int col = 3; col <= 35; col++) {
                Cell cell = dataRow.getCell(col);
                if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() != 0.0) {
                    hasData = true;
                    break;
                }
            }
            assertTrue(hasData, "Company 29 should have non-zero data in Sheet1");

            // AK column should have SUM formula
            Cell totalCell = dataRow.getCell(36);
            assertNotNull(totalCell);
            assertEquals(CellType.FORMULA, totalCell.getCellType());
            assertTrue(totalCell.getCellFormula().startsWith("SUM("));

            // Subtotal row
            Row subtotalRow = sheet1.getRow(3);
            assertNotNull(subtotalRow);
            assertEquals("小計", subtotalRow.getCell(2).getStringCellValue());
            Cell subtotalD = subtotalRow.getCell(3);
            assertEquals(CellType.FORMULA, subtotalD.getCellType());
            assertTrue(subtotalD.getCellFormula().contains("SUBTOTAL"));

            // === Sheet2 驗證 ===
            XSSFSheet sheet2 = workbook.getSheetAt(1);

            // Title
            String title = sheet2.getRow(0).getCell(0).getStringCellValue();
            assertTrue(title.contains("115年度第1季"), "Title should contain year and quarter");
            assertTrue(title.contains("1-3月份"), "Title should contain month range");

            // Year headers
            assertEquals("115年度", sheet2.getRow(3).getCell(19).getStringCellValue());
            assertEquals("114年度", sheet2.getRow(3).getCell(20).getStringCellValue());

            // Data row in Sheet2 (row 5, 0-based)
            Row s2DataRow = sheet2.getRow(5);
            assertNotNull(s2DataRow);
            assertEquals("29", s2DataRow.getCell(0).getStringCellValue());

            // Cross-sheet formula in D column
            Cell s2D = s2DataRow.getCell(3);
            assertNotNull(s2D);
            assertEquals(CellType.FORMULA, s2D.getCellType());
            assertTrue(s2D.getCellFormula().contains("115自留(季)"),
                    "Should have cross-sheet formula referencing Sheet1");

            // T column: SUM formula
            Cell s2T = s2DataRow.getCell(19);
            assertNotNull(s2T);
            assertEquals(CellType.FORMULA, s2T.getCellType());

            // U column: last year data
            Cell s2U = s2DataRow.getCell(20);
            assertNotNull(s2U);
            assertEquals(5000000.0, s2U.getNumericCellValue(), 0.01,
                    "U column should have last year data");

            // V column: growth formula
            Cell s2V = s2DataRow.getCell(21);
            assertNotNull(s2V);
            assertEquals(CellType.FORMULA, s2V.getCellType());

            // === Sheet3 (歸屬) 驗證 ===
            XSSFSheet sheet3 = workbook.getSheetAt(2);
            Row guishuHeader = sheet3.getRow(0);
            assertEquals("類", guishuHeader.getCell(0).getStringCellValue());
            assertEquals("歸屬", guishuHeader.getCell(1).getStringCellValue());
            assertEquals("代號", guishuHeader.getCell(2).getStringCellValue());
            assertEquals("險種", guishuHeader.getCell(4).getStringCellValue());

            Row guishuRow1 = sheet3.getRow(1);
            assertEquals("一", guishuRow1.getCell(0).getStringCellValue());
            assertEquals("火險", guishuRow1.getCell(1).getStringCellValue());
            assertEquals("0100", guishuRow1.getCell(2).getStringCellValue());
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

        String outputPath = tempDir.resolve("output_no_lastyear.xlsx").toString();

        reportWriter.writeReport(outputPath, quarterDataMap, lastYearData, 115, 1);

        try (FileInputStream fis = new FileInputStream(outputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet2 = workbook.getSheetAt(1);
            Row dataRow = sheet2.getRow(5);
            assertNotNull(dataRow);

            // U column should be empty (string "") when no last year data
            Cell uCell = dataRow.getCell(20);
            assertNotNull(uCell, "U column cell should exist");
        }
    }

    @Test
    void writeReport_withDifferentYear_shouldUpdateSheetNamesAndTitle() throws Exception {
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

        String outputPath = tempDir.resolve("output_116.xlsx").toString();

        reportWriter.writeReport(outputPath, quarterDataMap, lastYearData, 116, 1);

        try (FileInputStream fis = new FileInputStream(outputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            assertEquals("116自留(季)", workbook.getSheetName(0));
            assertEquals("116自留總累", workbook.getSheetName(1));
            assertEquals("歸屬", workbook.getSheetName(2));

            XSSFSheet sheet2 = workbook.getSheetAt(1);
            String title = sheet2.getRow(0).getCell(0).getStringCellValue();
            assertTrue(title.startsWith("116年度第1季"), "Title should use year 116");

            assertEquals("116年度", sheet2.getRow(3).getCell(19).getStringCellValue());
            assertEquals("115年度", sheet2.getRow(3).getCell(20).getStringCellValue());
        }
    }
}
