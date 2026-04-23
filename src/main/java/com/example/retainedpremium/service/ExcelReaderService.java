package com.example.retainedpremium.service;

import com.example.retainedpremium.model.CompanyData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import static com.example.retainedpremium.constant.InsuranceConstants.*;

@Service
public class ExcelReaderService {

    private static final Logger log = LoggerFactory.getLogger(ExcelReaderService.class);

    public Optional<CompanyData> readSourceFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            // Read company code from B4 (1-based row=4, col=2 → 0-based row=3, col=1)
            String companyCode = getCellStringValue(
                    sheet.getRow(SOURCE_COMPANY_CODE_ROW - 1).getCell(SOURCE_COMPANY_CODE_COL - 1));
            // Read company name from C4 (1-based row=4, col=3 → 0-based row=3, col=2)
            String companyName = getCellStringValue(
                    sheet.getRow(SOURCE_COMPANY_NAME_ROW - 1).getCell(SOURCE_COMPANY_NAME_COL - 1));

            Map<String, Double> retainedPremiums = new LinkedHashMap<>();

            // Iterate rows 7-39 (1-based) → POI row index 6-38
            for (int rowIdx = SOURCE_DATA_START_ROW - 1; rowIdx <= SOURCE_DATA_END_ROW - 1; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) continue;

                Cell codeCell = row.getCell(SOURCE_CODE_COL - 1);
                if (codeCell == null) continue;

                String insuranceCode = getCellStringValue(codeCell);
                if (insuranceCode == null || insuranceCode.isBlank()) continue;

                // Read retained premium from column F (formula cell)
                Cell retainedCell = row.getCell(SOURCE_RETAINED_COL - 1);
                double retainedPremium = evaluateNumericCell(retainedCell, evaluator);

                retainedPremiums.put(insuranceCode, retainedPremium);
            }

            log.info("Successfully read source file: {} - company={} ({}), entries={}",
                    filePath, companyCode, companyName, retainedPremiums.size());

            return Optional.of(new CompanyData(companyCode, companyName, retainedPremiums));

        } catch (Exception e) {
            log.error("Failed to read source file: {} - {}", filePath, e.getMessage(), e);
            return Optional.empty();
        }
    }

    /**
     * Reads a cell value as String, handling both NUMERIC and STRING types.
     * For NUMERIC cells, zero-pads to 4 digits to preserve codes like "0100".
     */
    String getCellStringValue(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.format("%04d", (int) cell.getNumericCellValue());
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
            // Fall back to cached numeric value
            try {
                return cell.getNumericCellValue();
            } catch (Exception ex) {
                return 0.0;
            }
        }
    }
}
