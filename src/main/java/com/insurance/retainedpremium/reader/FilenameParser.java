package com.insurance.retainedpremium.reader;

import com.insurance.retainedpremium.model.FileInfo;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.nio.file.Path;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Component
public class FilenameParser {

    private static final Logger log = LoggerFactory.getLogger(FilenameParser.class);

    private static final Pattern FILENAME_PATTERN = Pattern.compile(
            "^(\\d{1,2})_(\\d{2,3})\\((\\d{2})-(\\d{2})\\)_自留保費統計表\\.xlsx$"
    );

    private static final Set<Integer> VALID_END_MONTHS = Set.of(3, 6, 9, 12);

    public Optional<FileInfo> parse(String filepath) {
        String filename = Path.of(filepath).getFileName().toString();
        Matcher matcher = FILENAME_PATTERN.matcher(filename);

        if (!matcher.matches()) {
            log.error("Filename does not match expected pattern: {}", filename);
            return Optional.empty();
        }

        String companyCode = matcher.group(1);
        int year = Integer.parseInt(matcher.group(2));
        int startMonth = Integer.parseInt(matcher.group(3));
        int endMonth = Integer.parseInt(matcher.group(4));

        if (startMonth < 1 || startMonth > 12 || endMonth < 1 || endMonth > 12) {
            log.error("Invalid month values in filename: {} (startMonth={}, endMonth={})", filename, startMonth, endMonth);
            return Optional.empty();
        }

        if (startMonth > endMonth) {
            log.error("startMonth ({}) is greater than endMonth ({}) in filename: {}", startMonth, endMonth, filename);
            return Optional.empty();
        }

        if (startMonth != 1) {
            log.error("startMonth must be 01 for cumulative system, got {} in filename: {}", String.format("%02d", startMonth), filename);
            return Optional.empty();
        }

        if (!VALID_END_MONTHS.contains(endMonth)) {
            log.error("endMonth {} does not correspond to a valid quarter in filename: {}", String.format("%02d", endMonth), filename);
            return Optional.empty();
        }

        int quarter = endMonth / 3;

        return Optional.of(new FileInfo(companyCode, year, startMonth, endMonth, quarter, filename));
    }
}
