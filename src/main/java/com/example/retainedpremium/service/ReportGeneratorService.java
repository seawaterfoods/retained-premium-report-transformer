package com.example.retainedpremium.service;

import com.example.retainedpremium.model.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

@Service
public class ReportGeneratorService {
    private static final Logger log = LoggerFactory.getLogger(ReportGeneratorService.class);

    private final FilenameParserService filenameParser;
    private final ExcelReaderService excelReader;
    private final DataTransformerService dataTransformer;
    private final TemplateWriterService templateWriter;
    private final LastYearReaderService lastYearReader;
    private final ValidationService validator;

    public ReportGeneratorService(FilenameParserService filenameParser,
                                  ExcelReaderService excelReader,
                                  DataTransformerService dataTransformer,
                                  TemplateWriterService templateWriter,
                                  LastYearReaderService lastYearReader,
                                  ValidationService validator) {
        this.filenameParser = filenameParser;
        this.excelReader = excelReader;
        this.dataTransformer = dataTransformer;
        this.templateWriter = templateWriter;
        this.lastYearReader = lastYearReader;
        this.validator = validator;
    }

    public void generate(String inputDir, String templatePath, String outputDir, String lastYearDir) {
        // STEP 1: Scan input directory for .xlsx files
        log.info("掃描輸入目錄: {}", inputDir);
        List<File> xlsxFiles = scanInputFiles(inputDir);
        if (xlsxFiles.isEmpty()) {
            log.error("輸入目錄中沒有找到 .xlsx 檔案: {}", inputDir);
            return;
        }
        log.info("找到 {} 個 .xlsx 檔案", xlsxFiles.size());

        // STEP 2: Parse all filenames
        log.info("解析檔名...");
        List<FileInfo> fileInfos = new ArrayList<>();
        for (File f : xlsxFiles) {
            Optional<FileInfo> parsed = filenameParser.parse(f.getAbsolutePath());
            if (parsed.isPresent()) {
                fileInfos.add(parsed.get());
                log.info("  已解析: {} → 公司={}, 年度={}, Q{}", f.getName(),
                    parsed.get().companyCode(), parsed.get().year(), parsed.get().quarter());
            }
        }
        if (fileInfos.isEmpty()) {
            log.error("沒有有效的來源檔案");
            return;
        }

        // STEP 3: Pre-validation (year consistency, duplicates)
        log.info("前置驗證...");
        ValidationResult preValidation = validator.validateFileInfos(fileInfos);
        if (!preValidation.isPassed()) {
            log.error("前置驗證失敗，中止處理:");
            preValidation.getErrors().forEach(e -> log.error("  {}", e));
            return;
        }
        preValidation.getWarnings().forEach(w -> log.warn("  {}", w));

        // STEP 4: Read all source files
        log.info("讀取來源檔案...");
        Map<String, CompanyData> companyDataMap = new LinkedHashMap<>();
        boolean hasReadError = false;
        for (FileInfo fi : fileInfos) {
            File sourceFile = xlsxFiles.stream()
                .filter(f -> f.getName().equals(fi.filename()))
                .findFirst().orElse(null);
            if (sourceFile == null) {
                log.error("找不到對應檔案: {}", fi.filename());
                hasReadError = true;
                continue;
            }
            Optional<CompanyData> data = excelReader.readSourceFile(sourceFile.getAbsolutePath());
            if (data.isPresent()) {
                companyDataMap.put(fi.companyCode(), data.get());
                log.info("  已讀取: {} ({})", fi.companyCode(), data.get().companyName());
            } else {
                log.error("讀取失敗: {}", fi.filename());
                hasReadError = true;
            }
        }
        if (companyDataMap.isEmpty()) {
            log.error("所有來源檔案讀取失敗，中止處理");
            return;
        }
        if (hasReadError) {
            log.error("部分來源資料讀取失敗，中止處理");
            return;
        }

        // STEP 5: Validate company data
        ValidationResult dataValidation = validator.validateAll(fileInfos, companyDataMap);
        if (!dataValidation.isPassed()) {
            log.error("資料驗證失敗，中止處理:");
            dataValidation.getErrors().forEach(e -> log.error("  {}", e));
            return;
        }

        // STEP 6: Group by quarter
        log.info("按季度分組...");
        Map<Integer, QuarterData> quarterDataMap = dataTransformer.groupByQuarter(fileInfos, companyDataMap);
        int maxQuarter = quarterDataMap.keySet().stream().mapToInt(Integer::intValue).max().orElse(1);
        int year = fileInfos.get(0).year();
        log.info("年度={}, 最大季度=Q{}, 涵蓋季度={}", year, maxQuarter, quarterDataMap.keySet());

        // STEP 7: Read last year data
        log.info("讀取去年同期資料...");
        Map<String, Double> lastYearData = lastYearReader.readLastYearData(year, maxQuarter, lastYearDir);
        if (lastYearData.isEmpty()) {
            log.warn("無去年同期資料，U欄將留空");
        } else {
            log.info("已讀取去年資料，共 {} 家公司", lastYearData.size());
        }

        // STEP 8: Verify template exists
        if (!new File(templatePath).exists()) {
            log.error("模板檔案不存在: {}", templatePath);
            return;
        }

        // STEP 9: Generate output filename and path
        String outputFilename = String.format("%d年產險業務(Q%d季自留)保費統計表.xlsx", year, maxQuarter);
        Path outputPath = Paths.get(outputDir, outputFilename);

        // Ensure output directory exists
        try {
            Files.createDirectories(Paths.get(outputDir));
        } catch (Exception e) {
            log.error("無法建立輸出目錄: {}", outputDir, e);
            return;
        }

        // STEP 10: Write report
        log.info("寫入報表: {}", outputPath);
        try {
            templateWriter.writeReport(
                templatePath, outputPath.toString(),
                quarterDataMap, lastYearData, year, maxQuarter
            );
            log.info("報表產出完成: {}", outputPath);
        } catch (Exception e) {
            log.error("報表寫入失敗", e);
            return;
        }

        log.info("===== 處理完成 =====");
    }

    private List<File> scanInputFiles(String inputDir) {
        File dir = new File(inputDir);
        if (!dir.exists() || !dir.isDirectory()) {
            return Collections.emptyList();
        }
        File[] files = dir.listFiles((d, name) -> name.endsWith(".xlsx"));
        return files != null ? Arrays.asList(files) : Collections.emptyList();
    }
}
