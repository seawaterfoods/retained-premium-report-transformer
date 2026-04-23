package com.example.retainedpremium.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@ConfigurationProperties(prefix = "app")
public class AppConfig {
    private String inputDir;
    private String outputDir;
    private String templatePath;
    private String lastYearDir;

    public String getInputDir() { return inputDir; }
    public void setInputDir(String inputDir) { this.inputDir = inputDir; }
    public String getOutputDir() { return outputDir; }
    public void setOutputDir(String outputDir) { this.outputDir = outputDir; }
    public String getTemplatePath() { return templatePath; }
    public void setTemplatePath(String templatePath) { this.templatePath = templatePath; }
    public String getLastYearDir() { return lastYearDir; }
    public void setLastYearDir(String lastYearDir) { this.lastYearDir = lastYearDir; }
}
