package com.insurance.retainedpremium.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@ConfigurationProperties(prefix = "app")
public class AppConfig {
    private String importDir;
    private String outputDir;
    private String lastYearDir;

    public String getImportDir() { return importDir; }
    public void setImportDir(String importDir) { this.importDir = importDir; }
    public String getOutputDir() { return outputDir; }
    public void setOutputDir(String outputDir) { this.outputDir = outputDir; }
    public String getLastYearDir() { return lastYearDir; }
    public void setLastYearDir(String lastYearDir) { this.lastYearDir = lastYearDir; }
}
