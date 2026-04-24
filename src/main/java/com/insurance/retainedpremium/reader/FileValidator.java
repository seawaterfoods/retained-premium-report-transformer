package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.model.CompanyData;
import com.insurance.retainedpremium.model.FileInfo;
import com.insurance.retainedpremium.model.ValidationResult;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Component
public class FileValidator {

    private static final Logger log = LoggerFactory.getLogger(FileValidator.class);

    public ValidationResult validateFileInfos(List<FileInfo> fileInfos) {
        ValidationResult result = new ValidationResult();

        if (fileInfos == null || fileInfos.isEmpty()) {
            String msg = "No files to process";
            result.addError(msg);
            log.error(msg);
            return result;
        }

        // Year consistency
        int firstYear = fileInfos.get(0).year();
        for (FileInfo fi : fileInfos) {
            if (fi.year() != firstYear) {
                String msg = String.format("Inconsistent year in file '%s': expected %d but found %d",
                        fi.filename(), firstYear, fi.year());
                result.addError(msg);
                log.error(msg);
            }
        }

        // Duplicate company in same quarter
        Set<String> seen = new HashSet<>();
        for (FileInfo fi : fileInfos) {
            String key = fi.companyCode() + "_Q" + fi.quarter();
            if (!seen.add(key)) {
                String msg = String.format("Duplicate company '%s' in quarter %d (file: %s)",
                        fi.companyCode(), fi.quarter(), fi.filename());
                result.addError(msg);
                log.error(msg);
            }
        }

        return result;
    }

    public ValidationResult validateCompanyData(CompanyData data, String filename) {
        ValidationResult result = new ValidationResult();

        if (data.companyCode() == null || data.companyCode().isBlank()) {
            String msg = String.format("Company code is empty in file '%s'", filename);
            result.addError(msg);
            log.error(msg);
        }

        if (data.retainedPremiums() != null) {
            for (Map.Entry<String, Double> entry : data.retainedPremiums().entrySet()) {
                double value = entry.getValue();
                if (Double.isNaN(value) || Double.isInfinite(value)) {
                    String msg = String.format("Invalid numeric value for insurance code '%s' in file '%s': %s",
                            entry.getKey(), filename, value);
                    result.addError(msg);
                    log.error(msg);
                }
            }

            for (Map.Entry<String, Double> entry : data.retainedPremiums().entrySet()) {
                if (entry.getValue() < 0) {
                    String msg = String.format("Negative retained premium for insurance code '%s' in file '%s': %s",
                            entry.getKey(), filename, entry.getValue());
                    result.addWarning(msg);
                    log.warn(msg);
                }
            }

            if (data.retainedPremiums().size() != 33) {
                String msg = String.format("Expected 33 insurance codes but found %d in file '%s'",
                        data.retainedPremiums().size(), filename);
                result.addWarning(msg);
                log.warn(msg);
            }
        }

        return result;
    }

    public ValidationResult validateAll(List<FileInfo> fileInfos, Map<String, CompanyData> companyDataMap) {
        ValidationResult result = new ValidationResult();

        result.merge(validateFileInfos(fileInfos));

        for (Map.Entry<String, CompanyData> entry : companyDataMap.entrySet()) {
            result.merge(validateCompanyData(entry.getValue(), entry.getKey()));
        }

        log.info("Validation complete: {} errors, {} warnings",
                result.getErrors().size(), result.getWarnings().size());

        return result;
    }
}
