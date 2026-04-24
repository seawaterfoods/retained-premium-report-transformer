package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.model.FileInfo;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.Optional;

import static org.junit.jupiter.api.Assertions.*;

class FilenameParserTest {

    private FilenameParser parser;

    @BeforeEach
    void setUp() {
        parser = new FilenameParser();
    }

    @Test
    void shouldParseValidQ1Filename() {
        Optional<FileInfo> result = parser.parse("29_115(01-03)_自留保費統計表.xlsx");

        assertTrue(result.isPresent());
        FileInfo info = result.get();
        assertEquals("29", info.companyCode());
        assertEquals(115, info.year());
        assertEquals(1, info.startMonth());
        assertEquals(3, info.endMonth());
        assertEquals(1, info.quarter());
        assertEquals("29_115(01-03)_自留保費統計表.xlsx", info.filename());
    }

    @Test
    void shouldParseValidQ2Filename() {
        Optional<FileInfo> result = parser.parse("05_115(01-06)_自留保費統計表.xlsx");

        assertTrue(result.isPresent());
        FileInfo info = result.get();
        assertEquals("05", info.companyCode());
        assertEquals(115, info.year());
        assertEquals(1, info.startMonth());
        assertEquals(6, info.endMonth());
        assertEquals(2, info.quarter());
    }

    @Test
    void shouldParseValidQ4Filename() {
        Optional<FileInfo> result = parser.parse("01_115(01-12)_自留保費統計表.xlsx");

        assertTrue(result.isPresent());
        FileInfo info = result.get();
        assertEquals("01", info.companyCode());
        assertEquals(115, info.year());
        assertEquals(1, info.startMonth());
        assertEquals(12, info.endMonth());
        assertEquals(4, info.quarter());
    }

    @Test
    void shouldRejectWrongExtension() {
        Optional<FileInfo> result = parser.parse("29_115(01-03)_自留保費統計表.xls");
        assertTrue(result.isEmpty());
    }

    @Test
    void shouldRejectBadFormat() {
        Optional<FileInfo> result = parser.parse("invalid_file.xlsx");
        assertTrue(result.isEmpty());
    }

    @Test
    void shouldRejectStartMonthNotOne() {
        Optional<FileInfo> result = parser.parse("29_115(04-06)_自留保費統計表.xlsx");
        assertTrue(result.isEmpty());
    }

    @Test
    void shouldRejectInvalidEndMonth() {
        Optional<FileInfo> result = parser.parse("29_115(01-05)_自留保費統計表.xlsx");
        assertTrue(result.isEmpty());
    }

    @Test
    void shouldParseFilenameFromFullPath() {
        Optional<FileInfo> result = parser.parse("D:\\import\\29_115(01-03)_自留保費統計表.xlsx");

        assertTrue(result.isPresent());
        FileInfo info = result.get();
        assertEquals("29", info.companyCode());
        assertEquals(115, info.year());
        assertEquals(1, info.quarter());
        assertEquals("29_115(01-03)_自留保費統計表.xlsx", info.filename());
    }
}
