package com.example.retainedpremium.service;

import com.example.retainedpremium.constant.InsuranceConstants;
import com.example.retainedpremium.model.CompanyData;
import com.example.retainedpremium.model.QuarterData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

@Service
public class TemplateWriterService {

    private static final Logger log = LoggerFactory.getLogger(TemplateWriterService.class);

    private final DataTransformerService dataTransformerService;

    public TemplateWriterService(DataTransformerService dataTransformerService) {
        this.dataTransformerService = dataTransformerService;
    }

    public void writeReport(String templatePath, String outputPath,
                            Map<Integer, QuarterData> quarterDataMap,
                            Map<String, Double> lastYearData,
                            int year, int maxQuarter) throws IOException {

        try (FileInputStream fis = new FileInputStream(templatePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet1 = workbook.getSheetAt(0);
            XSSFSheet sheet2 = workbook.getSheetAt(1);

            // Write Sheet1 data and hide unused rows
            for (Map.Entry<Integer, int[]> blockEntry : InsuranceConstants.S1_QUARTER_BLOCKS.entrySet()) {
                int quarter = blockEntry.getKey();
                int[] block = blockEntry.getValue();
                int dataStart = block[0];
                int dataEnd = block[1];

                Map<String, Integer> companyRowMap = buildCompanyRowMap(sheet1, dataStart, dataEnd);
                Set<String> companiesWithData = new HashSet<>();

                QuarterData qd = quarterDataMap.get(quarter);
                if (qd != null) {
                    for (CompanyData companyData : qd.companies().values()) {
                        String code = companyData.companyCode();
                        Integer rowNum = companyRowMap.get(code);
                        if (rowNum == null) {
                            log.error("公司代號 {} 不在模板中，跳過", code);
                            continue;
                        }
                        companiesWithData.add(code);
                        Row row = sheet1.getRow(rowNum - 1);
                        for (Map.Entry<String, Integer> colEntry : InsuranceConstants.TEMPLATE_S1_CODE_TO_COL.entrySet()) {
                            int col = colEntry.getValue();
                            double value = companyData.retainedPremiums().getOrDefault(colEntry.getKey(), 0.0);
                            Cell cell = row.getCell(col - 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            cell.setCellValue(value);
                        }
                    }
                }

                // Hide/show rows
                for (Map.Entry<String, Integer> entry : companyRowMap.entrySet()) {
                    Row row = sheet1.getRow(entry.getValue() - 1);
                    if (row != null) {
                        row.setZeroHeight(!companiesWithData.contains(entry.getKey()));
                    }
                }
            }

            // Write Sheet2 data and hide unused rows
            for (Map.Entry<Integer, int[]> blockEntry : InsuranceConstants.S2_QUARTER_BLOCKS.entrySet()) {
                int quarter = blockEntry.getKey();
                int[] block = blockEntry.getValue();
                int dataStart = block[0];
                int dataEnd = block[1];

                Map<String, Integer> companyRowMap = buildCompanyRowMap(sheet2, dataStart, dataEnd);
                Set<String> companiesWithData = new HashSet<>();

                QuarterData qd = quarterDataMap.get(quarter);
                if (qd != null) {
                    for (CompanyData companyData : qd.companies().values()) {
                        String code = companyData.companyCode();
                        Integer rowNum = companyRowMap.get(code);
                        if (rowNum == null) {
                            log.error("公司代號 {} 不在模板中，跳過", code);
                            continue;
                        }
                        companiesWithData.add(code);
                        Row row = sheet2.getRow(rowNum - 1);

                        // Write aggregated categories (D-S, cols 4-19)
                        Map<Integer, Double> aggregated = dataTransformerService.aggregateCategories(companyData);
                        for (Map.Entry<Integer, Double> catEntry : aggregated.entrySet()) {
                            int col = catEntry.getKey(); // 1-based
                            Cell cell = row.getCell(col - 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            cell.setCellValue(catEntry.getValue());
                        }

                        // Write U column (last year data)
                        Double lastYear = lastYearData.get(code);
                        if (lastYear != null) {
                            Cell uCell = row.getCell(InsuranceConstants.TEMPLATE_S2_LASTYEAR_COL - 1,
                                    Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            uCell.setCellValue(lastYear);
                        } else {
                            log.warn("公司代號 {} 無去年資料，U欄留空", code);
                        }
                    }
                }

                // Hide/show rows
                for (Map.Entry<String, Integer> entry : companyRowMap.entrySet()) {
                    Row row = sheet2.getRow(entry.getValue() - 1);
                    if (row != null) {
                        row.setZeroHeight(!companiesWithData.contains(entry.getKey()));
                    }
                }
            }

            // Update Sheet2 title (row 1 = A1)
            int[] months = InsuranceConstants.QUARTER_MONTH_MAP.get(maxQuarter);
            String newTitle = year + "年度第" + maxQuarter + "季(" + months[0] + "-" + months[1]
                    + "月份)各保險公司自留保費統計總表";
            sheet2.getRow(0).getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(newTitle);

            // Update Sheet2 row 4 col T and U headers
            sheet2.getRow(3).getCell(19, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                    .setCellValue(year + "年度");
            sheet2.getRow(3).getCell(20, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                    .setCellValue((year - 1) + "年度");

            // Rename sheets
            workbook.setSheetName(0, year + "自留(季)");
            workbook.setSheetName(1, year + "自留總累");

            // Force formula recalculation
            workbook.setForceFormulaRecalculation(true);

            // Save output
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
            }

            log.info("Report written successfully to {}", outputPath);
        }
    }

    private Map<String, Integer> buildCompanyRowMap(XSSFSheet sheet, int dataStart, int dataEnd) {
        Map<String, Integer> map = new HashMap<>();
        for (int rowNum = dataStart; rowNum <= dataEnd; rowNum++) {
            Row row = sheet.getRow(rowNum - 1); // convert to 0-based
            if (row == null) continue;
            Cell cell = row.getCell(0); // column A
            if (cell == null) continue;

            String companyCode;
            if (cell.getCellType() == CellType.NUMERIC) {
                companyCode = String.format("%02d", (int) cell.getNumericCellValue());
            } else if (cell.getCellType() == CellType.STRING) {
                companyCode = cell.getStringCellValue().trim();
            } else {
                continue;
            }

            if (!companyCode.isEmpty()) {
                map.put(companyCode, rowNum); // store as 1-based
            }
        }
        return map;
    }
}
