package com.insurance.retainedpremium.reader;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class LastYearReaderTest {

    private LastYearReader lastYearReader;

    @BeforeEach
    void setUp() {
        lastYearReader = new LastYearReader();
    }

    @Test
    void readLastYearData_nonExistentDirectory_returnsEmptyMap() {
        Map<String, Double> result = lastYearReader.readLastYearData(
                2024, 1, "Z:\\nonexistent\\directory");

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void readLastYearData_existingDirectoryButNoFile_returnsEmptyMap(@TempDir Path tempDir) {
        Map<String, Double> result = lastYearReader.readLastYearData(
                2024, 1, tempDir.toString());

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }
}
