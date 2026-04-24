package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.model.CompanyData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import static com.insurance.retainedpremium.constant.InsuranceConstants.*;

@Component
public class ExcelSourceReader {

    private static final Logger log = LoggerFactory.getLogger(ExcelSourceReader.class);

    public Optional<CompanyData> readSourceFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            String companyCode = getCellStringValue(
                    sheet.getRow(SOURCE_COMPANY_CODE_ROW - 1).getCell(SOURCE_COMPANY_CODE_COL - 1));
            String companyName = getCellStringValue(
                    sheet.getRow(SOURCE_COMPANY_NAME_ROW - 1).getCell(SOURCE_COMPANY_NAME_COL - 1));

            Map<String, Double> retainedPremiums = new LinkedHashMap<>();

            for (int rowIdx = SOURCE_DATA_START_ROW - 1; rowIdx <= SOURCE_DATA_END_ROW - 1; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) continue;

                Cell codeCell = row.getCell(SOURCE_CODE_COL - 1);
                if (codeCell == null) continue;

                String insuranceCode = getCellStringValue(codeCell);
                if (insuranceCode == null || insuranceCode.isBlank()) continue;

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
            try {
                return cell.getNumericCellValue();
            } catch (Exception ex) {
                return 0.0;
            }
        }
    }
}
