package com.example.reporting.model;

import java.util.List;

public class SheetSummary {
    private final String sheetName;
    private final int rowCount;
    private final List<String> columnHeaders;
    private final List<List<String>> sampleRows;

    public SheetSummary(String sheetName, int rowCount, List<String> columnHeaders, List<List<String>> sampleRows) {
        this.sheetName = sheetName;
        this.rowCount = rowCount;
        this.columnHeaders = List.copyOf(columnHeaders);
        this.sampleRows = List.copyOf(sampleRows);
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getRowCount() {
        return rowCount;
    }

    public List<String> getColumnHeaders() {
        return columnHeaders;
    }

    public List<List<String>> getSampleRows() {
        return sampleRows;
    }
}
