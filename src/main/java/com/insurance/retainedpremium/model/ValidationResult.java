package com.insurance.retainedpremium.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 驗證結果
 */
public class ValidationResult {
    private boolean passed = true;
    private final List<String> errors = new ArrayList<>();
    private final List<String> warnings = new ArrayList<>();

    public void addError(String message) {
        errors.add(message);
        passed = false;
    }

    public void addWarning(String message) {
        warnings.add(message);
    }

    public boolean isPassed() { return passed; }
    public List<String> getErrors() { return errors; }
    public List<String> getWarnings() { return warnings; }

    public void merge(ValidationResult other) {
        this.errors.addAll(other.errors);
        this.warnings.addAll(other.warnings);
        if (!other.passed) {
            this.passed = false;
        }
    }
}
