package com.insurance.retainedpremium.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.ArrayList;
import java.util.List;

@Configuration
@ConfigurationProperties(prefix = "app")
public class AppConfig {
    private String importDir;
    private String outputDir;
    private Integer processYear;
    private InsuranceConfig insurance = new InsuranceConfig();

    public String getImportDir() { return importDir; }
    public void setImportDir(String importDir) { this.importDir = importDir; }
    public String getOutputDir() { return outputDir; }
    public void setOutputDir(String outputDir) { this.outputDir = outputDir; }
    public Integer getProcessYear() { return processYear; }
    public void setProcessYear(Integer processYear) { this.processYear = processYear; }
    public InsuranceConfig getInsurance() { return insurance; }
    public void setInsurance(InsuranceConfig insurance) { this.insurance = insurance; }

    public static class InsuranceConfig {
        private List<CodeConfig> codes = new ArrayList<>();
        private List<MajorCategoryConfig> categories = new ArrayList<>();

        public List<CodeConfig> getCodes() { return codes; }
        public void setCodes(List<CodeConfig> codes) { this.codes = codes; }
        public List<MajorCategoryConfig> getCategories() { return categories; }
        public void setCategories(List<MajorCategoryConfig> categories) { this.categories = categories; }
    }

    public static class CodeConfig {
        private String code;
        private String shortName;
        private String fullName;

        public String getCode() { return code; }
        public void setCode(String code) { this.code = code; }
        public String getShortName() { return shortName; }
        public void setShortName(String shortName) { this.shortName = shortName; }
        public String getFullName() { return fullName; }
        public void setFullName(String fullName) { this.fullName = fullName; }
    }

    public static class MajorCategoryConfig {
        private String name;
        private String number;
        private boolean overseas;
        private List<SubCategoryConfig> subCategories = new ArrayList<>();

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public String getNumber() { return number; }
        public void setNumber(String number) { this.number = number; }
        public boolean isOverseas() { return overseas; }
        public void setOverseas(boolean overseas) { this.overseas = overseas; }
        public List<SubCategoryConfig> getSubCategories() { return subCategories; }
        public void setSubCategories(List<SubCategoryConfig> subCategories) { this.subCategories = subCategories; }
    }

    public static class SubCategoryConfig {
        private String name;
        private String headerGroup;
        private String headerLabel;
        private String subHeader;
        private List<String> codes = new ArrayList<>();

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public String getHeaderGroup() { return headerGroup; }
        public void setHeaderGroup(String headerGroup) { this.headerGroup = headerGroup; }
        public String getHeaderLabel() { return headerLabel; }
        public void setHeaderLabel(String headerLabel) { this.headerLabel = headerLabel; }
        public String getSubHeader() { return subHeader; }
        public void setSubHeader(String subHeader) { this.subHeader = subHeader; }
        public List<String> getCodes() { return codes; }
        public void setCodes(List<String> codes) { this.codes = codes; }
    }
}
