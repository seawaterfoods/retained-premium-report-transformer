package com.example.retainedpremium.service;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class LastYearReaderServiceTest {

    private LastYearReaderService lastYearReaderService;

    @BeforeEach
    void setUp() {
        lastYearReaderService = new LastYearReaderService();
    }

    @Test
    void readLastYearData_nonExistentDirectory_returnsEmptyMap() {
        Map<String, Double> result = lastYearReaderService.readLastYearData(
                2024, 1, "Z:\\nonexistent\\directory");

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void readLastYearData_existingDirectoryButNoFile_returnsEmptyMap(@TempDir Path tempDir) {
        Map<String, Double> result = lastYearReaderService.readLastYearData(
                2024, 1, tempDir.toString());

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }
}
