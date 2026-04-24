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
            "^(\\d{1,2})_(\\d{2,3})Q(\\d)_自留保費統計表\\.xlsx$"
    );

    private static final Set<Integer> VALID_QUARTERS = Set.of(1, 2, 3, 4);

    public Optional<FileInfo> parse(String filepath) {
        String filename = Path.of(filepath).getFileName().toString();
        Matcher matcher = FILENAME_PATTERN.matcher(filename);

        if (!matcher.matches()) {
            log.error("Filename does not match expected pattern: {}", filename);
            return Optional.empty();
        }

        String companyCode = matcher.group(1);
        int year = Integer.parseInt(matcher.group(2));
        int quarter = Integer.parseInt(matcher.group(3));

        if (!VALID_QUARTERS.contains(quarter)) {
            log.error("Invalid quarter value in filename: {} (quarter={})", filename, quarter);
            return Optional.empty();
        }

        int startMonth = 1;
        int endMonth = quarter * 3;

        return Optional.of(new FileInfo(companyCode, year, startMonth, endMonth, quarter, filename));
    }
}
