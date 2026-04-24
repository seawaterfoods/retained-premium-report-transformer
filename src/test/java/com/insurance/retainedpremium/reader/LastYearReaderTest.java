package com.insurance.retainedpremium.reader;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.nio.file.Path;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class LastYearReaderTest {

    @Autowired
    private LastYearReader lastYearReader;

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
