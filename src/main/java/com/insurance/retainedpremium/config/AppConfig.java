package com.insurance.retainedpremium.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@ConfigurationProperties(prefix = "app")
public class AppConfig {
    private String importDir;
    private String outputDir;
    private Integer processYear;

    public String getImportDir() { return importDir; }
    public void setImportDir(String importDir) { this.importDir = importDir; }
    public String getOutputDir() { return outputDir; }
    public void setOutputDir(String outputDir) { this.outputDir = outputDir; }
    public Integer getProcessYear() { return processYear; }
    public void setProcessYear(Integer processYear) { this.processYear = processYear; }
}
