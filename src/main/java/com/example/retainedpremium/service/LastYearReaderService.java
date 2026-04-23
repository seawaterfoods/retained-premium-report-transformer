package com.example.retainedpremium.service;

import com.example.retainedpremium.constant.InsuranceConstants;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.LinkedHashMap;
import java.util.Map;

@Service
public class LastYearReaderService {

    private static final Logger log = LoggerFactory.getLogger(LastYearReaderService.class);

    /**
     * Reads last year's output report to get "去年同期" data for the U column in Sheet2.
     *
     * @param currentYear  the current report year
     * @param quarter      the quarter (1-4)
     * @param lastYearDir  directory containing last year's report file
     * @return companyCode → last year's T column value for the given quarter
     */
    public Map<String, Double> readLastYearData(int currentYear, int quarter, String lastYearDir) {
        String filename = (currentYear - 1) + "年產險業務(Q" + quarter + "季自留)保費統計表.xlsx";
        Path filePath = Paths.get(lastYearDir, filename);

        if (!Files.exists(filePath)) {
            log.warn("去年報表不存在: {}, U欄將留空", filename);
            return Map.of();
        }

        Map<String, Double> result = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(1);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            int[] block = InsuranceConstants.S2_QUARTER_BLOCKS.get(quarter);
            int dataStart = block[0];
            int dataEnd = block[1];

            // T column: 1-based col 20 → 0-based col 19
            int tColIndex = InsuranceConstants.TEMPLATE_S2_YEAR_TOTAL_COL - 1;

            for (int rowNum = dataStart; rowNum <= dataEnd; rowNum++) {
                Row row = sheet.getRow(rowNum - 1);
                if (row == null) continue;

                String companyCode = getCompanyCode(row.getCell(0));
                if (companyCode == null || companyCode.isBlank()) continue;

                Cell valueCell = row.getCell(tColIndex);
                double value = evaluateNumericCell(valueCell, evaluator);

                result.put(companyCode, value);
            }

            log.info("Read last year data from {}: {} companies", filename, result.size());

        } catch (Exception e) {
            log.error("Failed to read last year report: {} - {}", filename, e.getMessage(), e);
            return Map.of();
        }

        return result;
    }

    /**
     * Reads company code from column A, handling both NUMERIC and STRING cell types.
     * Numeric values are zero-padded to 2 digits.
     */
    private String getCompanyCode(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.format("%02d", (int) cell.getNumericCellValue());
            default -> null;
        };
    }

    private double evaluateNumericCell(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return 0.0;
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                evaluator.evaluateFormulaCell(cell);
            }
            return cell.getNumericCellValue();
        } catch (Exception e) {
            try {
                return cell.getNumericCellValue();
            } catch (Exception ex) {
                return 0.0;
            }
        }
    }
}
