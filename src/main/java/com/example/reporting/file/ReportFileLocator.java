package com.example.reporting.file;

import com.example.reporting.model.ReportFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ReportFileLocator {
    private static final Logger LOGGER = LoggerFactory.getLogger(ReportFileLocator.class);

    private final Path reportDirectory;
    private final Pattern filePattern;
    public ReportFileLocator(Path reportDirectory, String filePattern) {
        this.reportDirectory = reportDirectory;
        this.filePattern = compilePattern(filePattern);
    }

    public List<ReportFile> findReportFiles() {
        if (!Files.exists(reportDirectory)) {
            LOGGER.warn("Report directory does not exist: {}", reportDirectory);
            return List.of();
        }

        try (Stream<Path> stream = Files.list(reportDirectory)) {
            return stream
                    .filter(path -> !Files.isDirectory(path))
                    .map(this::toReportFile)
                    .flatMap(Optional::stream)
                    .sorted(Comparator.comparing(ReportFile::reportDate))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            LOGGER.error("Failed to list report directory {}", reportDirectory, e);
            return List.of();
        }
    }

    public Optional<ReportFile> findLatestReportFile() {
        List<ReportFile> files = findReportFiles();
        if (files.isEmpty()) {
            return Optional.empty();
        }
        return Optional.of(files.get(files.size() - 1));
    }

    private Optional<ReportFile> toReportFile(Path path) {
        String fileName = path.getFileName().toString();
        Matcher matcher = filePattern.matcher(fileName);
        if (!matcher.matches()) {
            LOGGER.debug("Skipping file that does not match pattern: {}", fileName);
            return Optional.empty();
        }

        String dateValue = matcher.group("year") + matcher.group("month") + matcher.group("day");
        try {
            LocalDate reportDate = LocalDate.parse(dateValue, DateTimeFormatter.BASIC_ISO_DATE);
            return Optional.of(new ReportFile(path, reportDate));
        } catch (DateTimeParseException ex) {
            LOGGER.warn("Failed to parse report date from file name: {}", fileName, ex);
            return Optional.empty();
        }
    }

    private Pattern compilePattern(String configuredPattern) {
        String escaped = configuredPattern
                .replace(".", "\\.")
                .replace("yyyy", "(?<year>\\d{4})")
                .replace("MM", "(?<month>\\d{2})")
                .replace("dd", "(?<day>\\d{2})");
        return Pattern.compile(escaped);
    }
}
