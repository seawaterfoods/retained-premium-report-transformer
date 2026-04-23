package com.example.retainedpremium.runner;

import com.example.retainedpremium.config.AppConfig;
import com.example.retainedpremium.service.ReportGeneratorService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;

@Component
public class ReportGeneratorRunner implements CommandLineRunner {
    private static final Logger log = LoggerFactory.getLogger(ReportGeneratorRunner.class);

    private final ReportGeneratorService reportGenerator;
    private final AppConfig appConfig;

    public ReportGeneratorRunner(ReportGeneratorService reportGenerator, AppConfig appConfig) {
        this.reportGenerator = reportGenerator;
        this.appConfig = appConfig;
    }

    @Override
    public void run(String... args) throws Exception {
        log.info("===== 自留保費統計表報表轉換系統 =====");
        log.info("輸入目錄: {}", appConfig.getInputDir());
        log.info("模板路徑: {}", appConfig.getTemplatePath());
        log.info("輸出目錄: {}", appConfig.getOutputDir());
        log.info("去年目錄: {}", appConfig.getLastYearDir());

        // Allow command line args to override config
        String inputDir = args.length > 0 ? args[0] : appConfig.getInputDir();
        String templatePath = args.length > 1 ? args[1] : appConfig.getTemplatePath();
        String outputDir = args.length > 2 ? args[2] : appConfig.getOutputDir();
        String lastYearDir = args.length > 3 ? args[3] : appConfig.getLastYearDir();

        reportGenerator.generate(inputDir, templatePath, outputDir, lastYearDir);
    }
}
