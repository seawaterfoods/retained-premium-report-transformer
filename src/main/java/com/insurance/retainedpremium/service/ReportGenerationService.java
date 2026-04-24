package com.insurance.retainedpremium.service;

import com.insurance.retainedpremium.config.AppConfig;
import com.insurance.retainedpremium.model.*;
import com.insurance.retainedpremium.reader.ExcelSourceReader;
import com.insurance.retainedpremium.reader.FileValidator;
import com.insurance.retainedpremium.reader.FilenameParser;
import com.insurance.retainedpremium.reader.LastYearReader;
import com.insurance.retainedpremium.writer.ReportWriter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Service;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

@Service
public class ReportGenerationService implements CommandLineRunner {
    private static final Logger log = LoggerFactory.getLogger(ReportGenerationService.class);

    private final AppConfig appConfig;
    private final FilenameParser filenameParser;
    private final ExcelSourceReader excelReader;
    private final DataTransformerService dataTransformer;
    private final ReportWriter reportWriter;
    private final LastYearReader lastYearReader;
    private final FileValidator validator;

    public ReportGenerationService(AppConfig appConfig,
                                   FilenameParser filenameParser,
                                   ExcelSourceReader excelReader,
                                   DataTransformerService dataTransformer,
                                   ReportWriter reportWriter,
                                   LastYearReader lastYearReader,
                                   FileValidator validator) {
        this.appConfig = appConfig;
        this.filenameParser = filenameParser;
        this.excelReader = excelReader;
        this.dataTransformer = dataTransformer;
        this.reportWriter = reportWriter;
        this.lastYearReader = lastYearReader;
        this.validator = validator;
    }

    @Override
    public void run(String... args) throws Exception {
        log.info("===== 自留保費統計表報表轉換系統 =====");
        log.info("匯入目錄: {}", appConfig.getImportDir());
        log.info("輸出目錄: {}", appConfig.getOutputDir());
        log.info("處理年度: {}", appConfig.getProcessYear());

        execute();
    }

    public void execute() {
        String importDir = appConfig.getImportDir();
        String outputDir = appConfig.getOutputDir();
        int year = appConfig.getProcessYear();

        // STEP 1: 掃描 import/{year}/Q{quarter}/ 目錄中的 .xlsx 檔案
        log.info("掃描匯入目錄: {}/{}/", importDir, year);
        Map<Integer, List<File>> quarterFiles = scanQuarterFiles(importDir, year);
        if (quarterFiles.isEmpty()) {
            log.error("匯入目錄中沒有找到任何季度資料: {}/{}/", importDir, year);
            return;
        }

        List<File> allFiles = new ArrayList<>();
        quarterFiles.values().forEach(allFiles::addAll);
        log.info("找到 {} 個 .xlsx 檔案，涵蓋 {} 個季度", allFiles.size(), quarterFiles.size());

        // STEP 2: Parse all filenames
        log.info("解析檔名...");
        List<FileInfo> fileInfos = new ArrayList<>();
        for (File f : allFiles) {
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

        // STEP 3: Pre-validation
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
            File sourceFile = allFiles.stream()
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
        log.info("年度={}, 最大季度=Q{}, 涵蓋季度={}", year, maxQuarter, quarterDataMap.keySet());

        // STEP 7: 讀取去年同期資料 (from output/{year-1}Q{quarter}/)
        log.info("讀取去年同期資料...");
        String lastYearOutputDir = Paths.get(outputDir, (year - 1) + "Q" + maxQuarter).toString();
        Map<String, Double> lastYearData = lastYearReader.readLastYearData(year, maxQuarter, lastYearOutputDir);
        if (lastYearData.isEmpty()) {
            log.warn("無去年同期資料，U欄將留空");
        } else {
            log.info("已讀取去年資料，共 {} 家公司", lastYearData.size());
        }

        // STEP 8: Generate output path: output/{year}Q{quarter}/
        String yearQuarterDir = year + "Q" + maxQuarter;
        Path outputSubDir = Paths.get(outputDir, yearQuarterDir);
        String outputFilename = String.format("%d年產險業務(Q%d季自留)保費統計表.xlsx", year, maxQuarter);
        Path outputPath = outputSubDir.resolve(outputFilename);

        // STEP 9: Write report
        log.info("寫入報表: {}", outputPath);
        try {
            reportWriter.writeReport(
                outputPath.toString(),
                quarterDataMap, lastYearData, year, maxQuarter
            );
            log.info("報表產出完成: {}", outputPath);
        } catch (Exception e) {
            log.error("報表寫入失敗", e);
            return;
        }

        log.info("===== 處理完成 =====");
    }

    /**
     * 掃描 import/{year}/Q1/ ~ Q4/ 目錄中的 .xlsx 檔案
     */
    private Map<Integer, List<File>> scanQuarterFiles(String importDir, int year) {
        Map<Integer, List<File>> result = new TreeMap<>();
        Path yearDir = Paths.get(importDir, String.valueOf(year));

        if (!Files.exists(yearDir) || !Files.isDirectory(yearDir)) {
            log.warn("年度目錄不存在: {}", yearDir);
            return result;
        }

        for (int q = 1; q <= 4; q++) {
            Path quarterDir = yearDir.resolve("Q" + q);
            if (Files.exists(quarterDir) && Files.isDirectory(quarterDir)) {
                File[] files = quarterDir.toFile().listFiles((d, name) -> name.endsWith(".xlsx"));
                if (files != null && files.length > 0) {
                    result.put(q, Arrays.asList(files));
                    log.info("  Q{}: {} 個檔案", q, files.length);
                }
            }
        }

        return result;
    }
}
