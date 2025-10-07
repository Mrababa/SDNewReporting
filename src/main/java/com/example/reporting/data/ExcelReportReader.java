package com.example.reporting.data;

import com.example.reporting.model.ReportFile;
import com.example.reporting.model.ReportSummary;
import com.example.reporting.model.SheetSummary;

import java.io.IOException;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelReportReader {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelReportReader.class);
    private static final int SAMPLE_ROW_LIMIT = 5;

    private final List<String> expectedSheets;

    public ExcelReportReader(List<String> expectedSheets) {
        this.expectedSheets = List.copyOf(expectedSheets);
    }

    public ReportSummary readReport(ReportFile reportFile) {
        Map<String, SheetSummary> sheetSummaries = new LinkedHashMap<>();
        List<String> missingSheets = new ArrayList<>();
        Path path = reportFile.path();

        try (Workbook workbook = WorkbookFactory.create(path.toFile())) {
            for (String sheetName : expectedSheets) {
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    missingSheets.add(sheetName);
                    LOGGER.warn("Missing expected sheet '{}' in report {}", sheetName, path.getFileName());
                    continue;
                }
                sheetSummaries.put(sheetName, summarizeSheet(sheet));
            }
        } catch (IOException e) {
            throw new IllegalStateException("Failed to read workbook: " + path, e);
        }

        return new ReportSummary(reportFile.reportDate(), sheetSummaries, missingSheets, OffsetDateTime.now());
    }

    private SheetSummary summarizeSheet(Sheet sheet) {
        DataFormatter formatter = new DataFormatter();
        List<String> headers = extractHeaders(sheet, formatter);
        List<List<String>> sampleRows = new ArrayList<>();
        int rowCount = 0;

        int firstRow = sheet.getFirstRowNum() + 1; // skip header
        int lastRow = sheet.getLastRowNum();
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null || isRowEmpty(row, formatter)) {
                continue;
            }
            rowCount++;
            if (sampleRows.size() < SAMPLE_ROW_LIMIT) {
                sampleRows.add(extractRow(row, headers.size(), formatter));
            }
        }

        return new SheetSummary(sheet.getSheetName(), rowCount, headers, sampleRows);
    }

    private List<String> extractHeaders(Sheet sheet, DataFormatter formatter) {
        Row headerRow = sheet.getRow(sheet.getFirstRowNum());
        List<String> headers = new ArrayList<>();
        if (headerRow == null) {
            return headers;
        }
        int lastCell = headerRow.getLastCellNum();
        for (int cellIndex = 0; cellIndex < lastCell; cellIndex++) {
            headers.add(formatter.formatCellValue(headerRow.getCell(cellIndex))); 
        }
        return headers;
    }

    private List<String> extractRow(Row row, int expectedCellCount, DataFormatter formatter) {
        int cellCount = Math.max(expectedCellCount, row.getLastCellNum());
        List<String> values = new ArrayList<>(cellCount);
        for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
            values.add(formatter.formatCellValue(row.getCell(cellIndex)));
        }
        return values;
    }

    private boolean isRowEmpty(Row row, DataFormatter formatter) {
        int lastCell = row.getLastCellNum();
        for (int cellIndex = 0; cellIndex < lastCell; cellIndex++) {
            String value = formatter.formatCellValue(row.getCell(cellIndex));
            if (!value.isBlank()) {
                return false;
            }
        }
        return true;
    }
}
