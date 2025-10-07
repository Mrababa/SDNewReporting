package com.example.reporting.model;

import java.time.LocalDate;
import java.time.OffsetDateTime;
import java.util.List;
import java.util.Map;

public class ReportSummary {
    private final LocalDate reportDate;
    private final Map<String, SheetSummary> sheetSummaries;
    private final List<String> missingSheets;
    private final OffsetDateTime generatedAt;

    public ReportSummary(LocalDate reportDate,
                         Map<String, SheetSummary> sheetSummaries,
                         List<String> missingSheets,
                         OffsetDateTime generatedAt) {
        this.reportDate = reportDate;
        this.sheetSummaries = Map.copyOf(sheetSummaries);
        this.missingSheets = List.copyOf(missingSheets);
        this.generatedAt = generatedAt;
    }

    public LocalDate getReportDate() {
        return reportDate;
    }

    public Map<String, SheetSummary> getSheetSummaries() {
        return sheetSummaries;
    }

    public List<String> getMissingSheets() {
        return missingSheets;
    }

    public OffsetDateTime getGeneratedAt() {
        return generatedAt;
    }

    public int getTotalRowCount() {
        return sheetSummaries.values().stream()
                .mapToInt(SheetSummary::getRowCount)
                .sum();
    }
}
