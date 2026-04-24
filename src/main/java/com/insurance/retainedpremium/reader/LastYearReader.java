package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.config.InsuranceMappingService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.LinkedHashMap;
import java.util.Map;

@Component
public class LastYearReader {

    private static final Logger log = LoggerFactory.getLogger(LastYearReader.class);

    private final InsuranceMappingService mapping;

    public LastYearReader(InsuranceMappingService mapping) {
        this.mapping = mapping;
    }

    /**
     * 讀取去年同期報表的 Sheet2 T 欄 (年度合計)。
     * 從 lastYearOutputDir 目錄中尋找去年報表。
     * 動態掃描所有列，以 A 欄公司代號 (純數字) 識別資料列。
     */
    public Map<String, Double> readLastYearData(int currentYear, int quarter, String lastYearOutputDir) {
        String filename = (currentYear - 1) + "年產險業務(Q" + quarter + "季自留)保費統計表.xlsx";
        Path filePath = Paths.get(lastYearOutputDir, filename);

        if (!Files.exists(filePath)) {
            log.warn("去年報表不存在: {}, U欄將留空", filename);
            return Map.of();
        }

        Map<String, Double> result = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(1);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            int tColIndex = mapping.getS2ColYearTotal() - 1;

            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;

                String companyCode = getCompanyCode(row.getCell(0));
                if (companyCode == null || companyCode.isBlank()) continue;
                if (!companyCode.matches("\\d+")) continue;

                Cell valueCell = row.getCell(tColIndex);
                double value = evaluateNumericCell(valueCell, evaluator);

                // 取最後一個季度區塊的值 (覆寫先前的)
                result.put(companyCode, value);
            }

            log.info("Read last year data from {}: {} companies", filename, result.size());

        } catch (Exception e) {
            log.error("Failed to read last year report: {} - {}", filename, e.getMessage(), e);
            return Map.of();
        }

        return result;
    }

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
