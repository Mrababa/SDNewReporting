package com.example.reporting.report;

import com.example.reporting.model.ReportSummary;
import com.example.reporting.model.SheetSummary;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import com.openhtmltopdf.pdfboxout.PdfRendererBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ReportGenerator {
    private static final Logger LOGGER = LoggerFactory.getLogger(ReportGenerator.class);
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("MMMM dd, yyyy", Locale.US);

    private static final List<String> ABNORMAL_IDS_COLUMNS = List.of("InsuranceCompanyName", "PolicyCount");
    private static final List<String> ICP_API_STATS_COLUMNS = List.of("ServiceName", "SuccessCount", "FailureCount");
    private static final List<String> ICP_ERRORS_COLUMNS = List.of(
            "ServiceName",
            "Error_Description",
            "ResponseCode",
            "Error_Count"
    );
    private static final List<String> MEM_UPLOAD_COLUMNS = List.of(
            "InsuranceCompanyName",
            "API_Upload",
            "Manual_Upload"
    );
    private static final List<String> SD_ERROR_COLUMNS = List.of(
            "InsuranceCompanyName",
            "FailureCount",
            "Total_API_Calls",
            "API_Failure_Ratio"
    );
    private static final List<String> SD_ERROR_DETAIL_COLUMNS = List.of(
            "InsuranceCompanyName",
            "ServiceName",
            "Error_Code",
            "Error_Description",
            "Error_Count"
    );
    private static final String[][] REPORT_SECTIONS = {
            {"abnormal-ids", "Abnormal Member IDs"},
            {"icp-api-stats", "ICP Service Success vs Failure"},
            {"icp-error-details", "ICP Failure Reasons"},
            {"mem-upload-counts", "Policy Upload Channels"},
            {"sd-error-ratio", "Service Failure Ratios"},
            {"sd-error-details-ic", "Error Details by Insurance Company"}
    };

    public ReportGenerator() {
    }

    public String buildHtmlReport(ReportSummary summary) {
        SheetSummary abnormalIds = summary.getSheetSummaries().get("VW_Abnormal_IDs");
        SheetSummary icpApiStats = summary.getSheetSummaries().get("VW_ICP_ApiSe_Stats");
        SheetSummary icpErrorDetails = summary.getSheetSummaries().get("VW_ICPSeErrorsDetails");
        SheetSummary memUploadCounts = summary.getSheetSummaries().get("VW_MemUploadTCount");
        SheetSummary sdServiceErrors = summary.getSheetSummaries().get("VW_SD_SeErrorDetails");
        SheetSummary sdServiceErrorsByIc = summary.getSheetSummaries().get("VW_SD_SeErrorDetailsIC");

        List<String> abnormalColumns = resolveColumns(abnormalIds, ABNORMAL_IDS_COLUMNS);
        List<String> icpApiColumns = resolveColumns(icpApiStats, ICP_API_STATS_COLUMNS);
        List<String> icpErrorColumns = resolveColumns(icpErrorDetails, ICP_ERRORS_COLUMNS);
        List<String> memUploadColumns = resolveColumns(memUploadCounts, MEM_UPLOAD_COLUMNS);
        List<String> sdErrorColumns = resolveColumns(sdServiceErrors, SD_ERROR_COLUMNS);
        List<String> sdErrorDetailColumns = resolveColumns(sdServiceErrorsByIc, SD_ERROR_DETAIL_COLUMNS);

        String abnormalJson = toJsonRows(abnormalIds, abnormalColumns);
        String icpApiJson = toJsonRows(icpApiStats, icpApiColumns);
        String icpErrorJson = toJsonRows(icpErrorDetails, icpErrorColumns);
        String memUploadJson = toJsonRows(memUploadCounts, memUploadColumns);
        String sdErrorJson = toJsonRows(sdServiceErrors, sdErrorColumns);
        String sdErrorDetailJson = toJsonRows(sdServiceErrorsByIc, sdErrorDetailColumns);

        StringBuilder html = new StringBuilder();
        html.append("<!DOCTYPE html>")
                .append("<html lang=\"en\">")
                .append("<head>")
                .append("<meta charset=\"UTF-8\" />")
                .append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />")
                .append("<title>Weekly SD Report</title>")
                .append("<link rel=\"stylesheet\" href=\"https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css\" />")
                .append("<link rel=\"stylesheet\" href=\"https://cdn.datatables.net/buttons/2.4.2/css/buttons.dataTables.min.css\" />")
                .append("<style>")
                .append(getStyles())
                .append("</style>")
                .append("</head>")
                .append("<body>")
                .append("<div class=\"page-shell\">");

        html.append("<header class=\"page-header\">")
                .append("<div class=\"branding\">")
                .append("<span class=\"page-badge\">Weekly Insight</span>")
                .append("<h1>SD Interactive Reporting Dashboard</h1>")
                .append("<p class=\"subtitle\">Operational quality metrics for the Service Desk platform.</p>")
                .append("</div>")
                .append("<div class=\"meta-grid\">")
                .append(String.format("<div class=\"meta-item\"><span class=\"meta-label\">Report Date</span><span class=\"meta-value\">%s</span></div>",
                        DATE_FORMATTER.format(summary.getReportDate())))
                .append(String.format("<div class=\"meta-item\"><span class=\"meta-label\">Generated At</span><span class=\"meta-value\">%s</span></div>",
                        escapeHtml(summary.getGeneratedAt().toString())))
                .append(String.format("<div class=\"meta-item\"><span class=\"meta-label\">Total Rows Processed</span><span class=\"meta-value\">%s</span></div>",
                        formatInteger(summary.getTotalRowCount())))
                .append("</div>")
                .append("<div class=\"callout\"><strong>Tip:</strong> Use the export buttons on each table to download Excel or PDF extracts for your working papers.</div>")
                .append("</header>");

        html.append("<main class=\"content\">")
                .append("<section class=\"summary-grid\">")
                .append(renderKeyMetrics(summary))
                .append(renderNavigation())
                .append(renderOverviewCard())
                .append("</section>");

        String missingSection = renderMissingSheets(summary);
        if (!missingSection.isEmpty()) {
            html.append(missingSection);
        }

        html.append(renderDataSection(
                "abnormal-ids",
                "Abnormal Member IDs",
                "Number of policies uploaded by each insurance company that contained incorrect or dummy identifiers.",
                "abnormalIdsChartContainer",
                "abnormalIdsChart",
                "abnormalIdsTable",
                "abnormalIdsEmpty"
        ));

        html.append(renderDataSection(
                "icp-api-stats",
                "ICP Service Success vs Failure",
                "Success and failure counts for the core ICP services.",
                "icpApiChartContainer",
                "icpApiChart",
                "icpApiTable",
                "icpApiEmpty"
        ));

        html.append(renderTableOnlySection(
                "icp-error-details",
                "ICP Failure Reasons",
                "Sortable and filterable list of failure reasons, including response codes and volumes.",
                "icpErrorsTable",
                "icpErrorsEmpty"
        ));

        html.append(renderDataSection(
                "mem-upload-counts",
                "Policy Upload Channels",
                "Comparison of API uploads vs manual uploads for each insurance company.",
                "memUploadChartContainer",
                "memUploadChart",
                "memUploadTable",
                "memUploadEmpty"
        ));

        html.append(renderDataSection(
                "sd-error-ratio",
                "Service Failure Ratios",
                "Failure counts, total API calls, and calculated failure ratios by insurance company.",
                "sdErrorChartContainer",
                "sdErrorChart",
                "sdErrorTable",
                "sdErrorEmpty"
        ));

        html.append(renderDataSection(
                "sd-error-details-ic",
                "Error Details by Insurance Company",
                "Detailed breakdown of failures by service, code, and description. Includes a chart of top recurring errors.",
                "sdErrorDetailChartContainer",
                "sdErrorDetailChart",
                "sdErrorDetailTable",
                "sdErrorDetailEmpty"
        ));

        html.append("</main>");

        html.append("<footer><p class=\"meta\">Report generated automatically from the latest workbook.</p></footer>")
                .append("</div>");

        html.append("<script src=\"https://code.jquery.com/jquery-3.7.1.min.js\"></script>")
                .append("<script src=\"https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js\"></script>")
                .append("<script src=\"https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js\"></script>")
                .append("<script src=\"https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js\"></script>")
                .append("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js\"></script>")
                .append("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js\"></script>")
                .append("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js\"></script>")
                .append("<script src=\"https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js\"></script>")
                .append("<script>")
                .append("const reportSections = ").append(toJsonSections()).append(";")
                .append("const abnormalIdsColumns = ").append(toJsonArray(abnormalColumns)).append(";")
                .append("const abnormalIdsData = ").append(abnormalJson).append(";")
                .append("const icpApiColumns = ").append(toJsonArray(icpApiColumns)).append(";")
                .append("const icpApiData = ").append(icpApiJson).append(";")
                .append("const icpErrorColumns = ").append(toJsonArray(icpErrorColumns)).append(";")
                .append("const icpErrorData = ").append(icpErrorJson).append(";")
                .append("const memUploadColumns = ").append(toJsonArray(memUploadColumns)).append(";")
                .append("const memUploadData = ").append(memUploadJson).append(";")
                .append("const sdErrorColumns = ").append(toJsonArray(sdErrorColumns)).append(";")
                .append("const sdErrorData = ").append(sdErrorJson).append(";")
                .append("const sdErrorDetailColumns = ").append(toJsonArray(sdErrorDetailColumns)).append(";")
                .append("const sdErrorDetailData = ").append(sdErrorDetailJson).append(";")
                .append(getScript())
                .append("</script>");

        html.append("</body></html>");
        return html.toString();
    }

    public Path writeHtmlReport(Path outputDirectory, String fileName, String htmlContent) {
        Path filePath = outputDirectory.resolve(fileName);
        try {
            Files.createDirectories(outputDirectory);
            Files.writeString(filePath, htmlContent, StandardCharsets.UTF_8);
            LOGGER.info("HTML report written to {}", filePath);
            return filePath;
        } catch (IOException e) {
            throw new UncheckedIOException("Failed to write HTML report", e);
        }
    }

    public Path writePdfReport(Path outputDirectory, String fileName, String htmlContent) {
        Path filePath = outputDirectory.resolve(fileName);
        try {
            Files.createDirectories(outputDirectory);
            try (var outputStream = Files.newOutputStream(filePath)) {
                PdfRendererBuilder builder = new PdfRendererBuilder();
                builder.useFastMode();
                builder.withHtmlContent(htmlContent, null);
                builder.toStream(outputStream);
                builder.run();
            }
            LOGGER.info("PDF report written to {}", filePath);
            return filePath;
        } catch (IOException e) {
            throw new UncheckedIOException("Failed to write PDF report", e);
        } catch (Exception e) {
            throw new IllegalStateException("Failed to generate PDF report", e);
        }
    }

    private List<String> resolveColumns(SheetSummary summary, List<String> preferredOrder) {
        if (summary == null) {
            return preferredOrder;
        }
        List<String> result = new ArrayList<>();
        if (preferredOrder != null) {
            for (String column : preferredOrder) {
                if (summary.getColumnHeaders().contains(column)) {
                    result.add(column);
                }
            }
        }
        for (String header : summary.getColumnHeaders()) {
            if (!result.contains(header)) {
                result.add(header);
            }
        }
        return result;
    }

    private String toJsonRows(SheetSummary summary, List<String> columns) {
        if (summary == null || summary.getRows().isEmpty()) {
            return "[]";
        }
        Map<String, Integer> columnIndex = new HashMap<>();
        List<String> headers = summary.getColumnHeaders();
        for (int i = 0; i < headers.size(); i++) {
            columnIndex.put(headers.get(i), i);
        }

        StringBuilder sb = new StringBuilder("[");
        boolean firstRow = true;
        for (List<String> row : summary.getRows()) {
            if (!firstRow) {
                sb.append(',');
            }
            firstRow = false;
            sb.append('{');
            boolean firstColumn = true;
            for (String column : columns) {
                if (!firstColumn) {
                    sb.append(',');
                }
                firstColumn = false;
                sb.append('"').append(escapeJson(column)).append('"').append(':');
                String value = "";
                Integer index = columnIndex.get(column);
                if (index != null && index < row.size()) {
                    value = row.get(index);
                }
                sb.append('"').append(escapeJson(value)).append('"');
            }
            sb.append('}');
        }
        sb.append(']');
        return sb.toString();
    }

    private String toJsonArray(List<String> values) {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i < values.size(); i++) {
            if (i > 0) {
                sb.append(',');
            }
            sb.append('"').append(escapeJson(values.get(i))).append('"');
        }
        sb.append(']');
        return sb.toString();
    }

    private String toJsonSections() {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i < REPORT_SECTIONS.length; i++) {
            if (i > 0) {
                sb.append(',');
            }
            sb.append('{')
                    .append("\"id\":\"").append(escapeJson(REPORT_SECTIONS[i][0])).append("\",")
                    .append("\"label\":\"").append(escapeJson(REPORT_SECTIONS[i][1])).append("\"")
                    .append('}');
        }
        sb.append(']');
        return sb.toString();
    }

    private String renderDataSection(String sectionId,
                                     String title,
                                     String description,
                                     String chartContainerId,
                                     String chartId,
                                     String tableId,
                                     String emptyMessageId) {
        String tabButtonId = "tab-" + sectionId;
        return new StringBuilder()
                .append('<').append("section class=\"card data-section tab-panel\" id=\"").append(sectionId)
                .append("\" role=\"tabpanel\" aria-labelledby=\"").append(tabButtonId).append("\" data-tab-panel=\"")
                .append(sectionId).append("\" tabindex=\"-1\">")
                .append("<div class=\"section-header\">")
                .append(String.format("<h2>%s</h2>", escapeHtml(title)))
                .append(String.format("<p class=\"description\">%s</p>", escapeHtml(description)))
                .append("</div>")
                .append(String.format("<div class=\"chart-card\" id=\"%s\"><canvas id=\"%s\"></canvas></div>",
                        chartContainerId,
                        chartId))
                .append(String.format("<p id=\"%s\" class=\"empty hidden\">No data available for this worksheet.</p>", emptyMessageId))
                .append(String.format("<div class=\"table-card\"><table id=\"%s\" class=\"display compact\" style=\"width:100%%\"></table></div>",
                        tableId))
                .append("</section>")
                .toString();
    }

    private String renderTableOnlySection(String sectionId,
                                          String title,
                                          String description,
                                          String tableId,
                                          String emptyMessageId) {
        String tabButtonId = "tab-" + sectionId;
        return new StringBuilder()
                .append('<').append("section class=\"card data-section tab-panel\" id=\"").append(sectionId)
                .append("\" role=\"tabpanel\" aria-labelledby=\"").append(tabButtonId).append("\" data-tab-panel=\"")
                .append(sectionId).append("\" tabindex=\"-1\">")
                .append("<div class=\"section-header\">")
                .append(String.format("<h2>%s</h2>", escapeHtml(title)))
                .append(String.format("<p class=\"description\">%s</p>", escapeHtml(description)))
                .append("</div>")
                .append(String.format("<p id=\"%s\" class=\"empty hidden\">No data available for this worksheet.</p>", emptyMessageId))
                .append(String.format("<div class=\"table-card\"><table id=\"%s\" class=\"display compact\" style=\"width:100%%\"></table></div>",
                        tableId))
                .append("</section>")
                .toString();
    }

    private String renderKeyMetrics(ReportSummary summary) {
        int processedSheets = summary.getSheetSummaries().size();
        int missingSheets = summary.getMissingSheets().size();
        int expectedSheets = processedSheets + missingSheets;
        int totalRows = summary.getTotalRowCount();

        SheetSummary topSheet = summary.getSheetSummaries().values().stream()
                .max(Comparator.comparingInt(SheetSummary::getRowCount))
                .orElse(null);

        String coverage = expectedSheets > 0
                ? String.format(Locale.US, "%.0f%%", (processedSheets * 100.0) / expectedSheets)
                : "—";
        String missingCaption = missingSheets > 0
                ? formatInteger(missingSheets) + " worksheet(s) outstanding"
                : "All expected worksheets present";
        String coverageCaption = expectedSheets > 0
                ? formatInteger(processedSheets) + " of " + formatInteger(expectedSheets) + " worksheets received"
                : "Awaiting initial data load";
        String recordsCaption = processedSheets > 0
                ? "Across " + formatInteger(processedSheets) + " worksheet(s)"
                : "No worksheet data available";
        String topSheetName = topSheet != null ? topSheet.getSheetName() : "—";
        String topSheetCaption = topSheet != null
                ? formatInteger(topSheet.getRowCount()) + " rows captured"
                : "No worksheet data available";

        return new StringBuilder()
                .append("<section class=\"card metrics\">")
                .append("<h2>Key Figures</h2>")
                .append("<div class=\"metrics-grid\">")
                .append(renderMetric("Worksheets processed", formatInteger(processedSheets), missingCaption))
                .append(renderMetric("Data coverage", coverage, coverageCaption))
                .append(renderMetric("Records analysed", formatInteger(totalRows), recordsCaption))
                .append(renderMetric("Top-volume worksheet", topSheetName, topSheetCaption))
                .append("</div>")
                .append("</section>")
                .toString();
    }

    private String renderMetric(String label, String primaryValue, String caption) {
        StringBuilder builder = new StringBuilder()
                .append("<div class=\"metric\">")
                .append(String.format("<span class=\"metric-label\">%s</span>", escapeHtml(label)))
                .append(String.format("<span class=\"metric-value\">%s</span>", escapeHtml(primaryValue == null ? "" : primaryValue)));
        if (caption != null && !caption.isBlank()) {
            builder.append(String.format("<span class=\"metric-caption\">%s</span>", escapeHtml(caption)));
        }
        builder.append("</div>");
        return builder.toString();
    }

    private String renderNavigation() {
        StringBuilder builder = new StringBuilder()
                .append("<section class=\"card quick-links\">")
                .append("<h2>Report Sections</h2>")
                .append("<div class=\"tab-bar\" role=\"tablist\">");
        for (String[] section : REPORT_SECTIONS) {
            builder.append(renderNavItem(section[0], section[1]));
        }
        builder.append("</div>")
                .append("<p class=\"tab-hint\">Switch between tabs to focus on one worksheet at a time.</p>")
                .append("</section>");
        return builder.toString();
    }

    private String renderNavItem(String id, String label) {
        String buttonId = "tab-" + id;
        return String.format(
                Locale.ROOT,
                "<button type=\"button\" class=\"tab-link\" role=\"tab\" id=\"%s\" data-tab=\"%s\" aria-controls=\"%s\" aria-selected=\"false\" tabindex=\"-1\">%s</button>",
                escapeHtml(buttonId),
                escapeHtml(id),
                escapeHtml(id),
                escapeHtml(label)
        );
    }

    private String renderOverviewCard() {
        return new StringBuilder()
                .append("<section class=\"card narrative\">")
                .append("<h2>How to interpret this dashboard</h2>")
                .append("<p>Use the curated visualisations and tables to monitor partner performance, surface outliers, and prepare executive-ready summaries.</p>")
                .append("<ul class=\"narrative-list\">")
                .append("<li><strong>Charts</strong> spotlight where performance is shifting; hover to reveal precise counts.</li>")
                .append("<li><strong>Filters</strong> and column sorting in each table accelerate deep dives without exporting the full dataset.</li>")
                .append("<li><strong>Exports</strong> (Excel/PDF) deliver ready-to-share extracts that mirror the on-screen filters.</li>")
                .append("</ul>")
                .append("</section>")
                .toString();
    }

    private String renderMissingSheets(ReportSummary summary) {
        if (summary.getMissingSheets().isEmpty()) {
            return "";
        }
        StringBuilder builder = new StringBuilder()
                .append("<section class=\"card missing\">")
                .append("<h2>Missing Worksheets</h2>")
                .append("<p>The following worksheets were not present in the source workbook and should be investigated:</p>")
                .append("<ul>");
        summary.getMissingSheets().forEach(sheet ->
                builder.append(String.format("<li>%s</li>", escapeHtml(sheet))));
        builder.append("</ul></section>");
        return builder.toString();
    }

    private String formatInteger(int value) {
        return String.format(Locale.US, "%,d", value);
    }

    private String escapeHtml(String value) {
        if (value == null) {
            return "";
        }
        return value.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#39;");
    }

    private String escapeJson(String value) {
        if (value == null) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            switch (c) {
                case '\\' -> sb.append("\\\\");
                case '"' -> sb.append("\\\"");
                case '\n' -> sb.append("\\n");
                case '\r' -> sb.append("\\r");
                case '\t' -> sb.append("\\t");
                case '<' -> sb.append("\\u003c");
                case '>' -> sb.append("\\u003e");
                case '&' -> sb.append("\\u0026");
                default -> {
                    if (c < 0x20) {
                        sb.append(String.format("\\u%04x", (int) c));
                    } else {
                        sb.append(c);
                    }
                }
            }
        }
        return sb.toString();
    }

    private String getStyles() {
        return "*, *::before, *::after { box-sizing: border-box; }"
                + "body { margin: 0; font-family: 'Segoe UI', Arial, sans-serif; background: #f1f5f9; color: #0f172a; }"
                + "a { color: #0ea5e9; text-decoration: none; }"
                + "a:hover { text-decoration: underline; }"
                + ".page-shell { max-width: 1200px; margin: 0 auto; padding: 2.5rem 2rem 3rem; }"
                + ".page-header { background: linear-gradient(135deg, #0f172a, #1e293b); color: #f8fafc; padding: 2rem; border-radius: 16px; box-shadow: 0 20px 35px -25px rgba(15, 23, 42, 0.7); margin-bottom: 2.5rem; display: grid; gap: 1.5rem; }"
                + ".branding { display: flex; flex-direction: column; gap: 0.5rem; }"
                + ".page-badge { align-self: flex-start; background: rgba(14, 165, 233, 0.25); color: #bae6fd; border-radius: 999px; padding: 0.25rem 0.85rem; font-size: 0.75rem; letter-spacing: 0.08em; text-transform: uppercase; }"
                + ".branding h1 { margin: 0; font-size: 2.1rem; font-weight: 600; }"
                + ".subtitle { margin: 0; font-size: 1.05rem; color: rgba(226, 232, 240, 0.85); }"
                + ".meta-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1rem; }"
                + ".meta-item { background: rgba(15, 23, 42, 0.4); border-radius: 12px; padding: 0.85rem 1rem; }"
                + ".meta-label { display: block; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.08em; color: rgba(226, 232, 240, 0.7); margin-bottom: 0.35rem; }"
                + ".meta-value { font-size: 1.15rem; font-weight: 600; color: #f8fafc; }"
                + ".callout { background: rgba(148, 163, 184, 0.18); border-radius: 12px; padding: 0.9rem 1.1rem; font-size: 0.95rem; color: #e2e8f0; border: 1px solid rgba(148, 163, 184, 0.35); }"
                + ".content { display: flex; flex-direction: column; gap: 2.5rem; }"
                + ".summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 1.5rem; }"
                + ".summary-grid .metrics { grid-column: span 2; }"
                + ".card { background: #ffffff; border-radius: 16px; padding: 1.75rem; box-shadow: 0 18px 40px -28px rgba(15, 23, 42, 0.55); border: 1px solid #e2e8f0; }"
                + ".card h2 { margin: 0; font-size: 1.4rem; font-weight: 600; color: #0f172a; }"
                + ".metrics-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1.25rem; margin-top: 1.5rem; }"
                + ".metric { display: flex; flex-direction: column; gap: 0.35rem; padding: 1rem 1.1rem; border-radius: 12px; background: linear-gradient(135deg, rgba(14, 165, 233, 0.08), rgba(99, 102, 241, 0.08)); border: 1px solid #e2e8f0; }"
                + ".metric-value { font-size: 1.6rem; font-weight: 600; color: #0f172a; }"
                + ".metric-label { font-size: 0.72rem; letter-spacing: 0.08em; text-transform: uppercase; color: #64748b; }"
                + ".metric-caption { font-size: 0.9rem; color: #475569; }"
                + ".quick-links { display: flex; flex-direction: column; gap: 1rem; }"
                + ".tab-bar { display: flex; flex-wrap: wrap; gap: 0.75rem; margin-top: 1.25rem; }"
                + ".tab-link { padding: 0.65rem 1rem; border-radius: 999px; border: 1px solid #cbd5f5; background: #f8fafc; color: #0f172a; font-weight: 500; cursor: pointer; transition: all 0.2s ease; }"
                + ".tab-link:hover { background: #e0f2fe; border-color: #0ea5e9; color: #0369a1; }"
                + ".tab-link.active { background: #0ea5e9; color: #f8fafc; border-color: #0ea5e9; box-shadow: 0 8px 18px -12px rgba(14, 165, 233, 0.8); }"
                + ".tab-link:focus-visible { outline: 3px solid rgba(59, 130, 246, 0.4); outline-offset: 2px; }"
                + ".tab-hint { margin: 0; font-size: 0.9rem; color: #475569; }"
                + ".narrative p { margin-top: 1rem; line-height: 1.6; color: #475569; }"
                + ".narrative-list { margin: 1.2rem 0 0; padding-left: 1.2rem; color: #475569; }"
                + ".narrative-list li { margin-bottom: 0.55rem; }"
                + ".missing { border-left: 4px solid #f97316; }"
                + ".missing h2 { color: #b91c1c; }"
                + ".missing p { margin-top: 0.75rem; color: #7f1d1d; }"
                + ".missing ul { margin: 1rem 0 0; padding-left: 1.3rem; color: #991b1b; }"
                + ".tab-panel { display: none; }"
                + ".tab-panel.active { display: grid; gap: 1.25rem; }"
                + ".section-header { display: flex; flex-direction: column; gap: 0.45rem; }"
                + ".description { margin: 0; color: #475569; line-height: 1.5; }"
                + ".chart-card { position: relative; min-height: 320px; padding: 1rem; background: #f8fafc; border-radius: 12px; border: 1px dashed #cbd5f5; }"
                + ".chart-card canvas { width: 100% !important; height: 100% !important; }"
                + ".table-card { overflow-x: auto; border-radius: 12px; border: 1px solid #e2e8f0; background: #ffffff; }"
                + "table.dataTable { width: 100% !important; border-collapse: collapse; font-size: 0.95rem; }"
                + "table.dataTable thead th { background: #e2e8f0; color: #0f172a; border-bottom: none; }"
                + "table.dataTable tbody tr:nth-child(even) { background: #f8fafc; }"
                + "table.dataTable tbody tr:hover { background: #e0f2fe; }"
                + ".empty { color: #64748b; font-style: italic; margin: 0; }"
                + ".hidden { display: none; }"
                + ".highlight-manual { background-color: #fef3c7 !important; }"
                + "footer { margin-top: 3rem; text-align: center; color: #475569; font-size: 0.85rem; }"
                + "@media (max-width: 900px) { .summary-grid .metrics { grid-column: span 1; } }"
                + "@media (max-width: 640px) { .page-shell { padding: 2rem 1.25rem 2.5rem; } }";
    }

    private String getScript() {
        StringBuilder script = new StringBuilder();
        script.append("const dataTables = {};\n");
        script.append("document.addEventListener('DOMContentLoaded', function () {\n");
        script.append("  initTabs(reportSections);\n");
        script.append("  initAbnormalIds();\n");
        script.append("  initIcpApiStats();\n");
        script.append("  initIcpErrorDetails();\n");
        script.append("  initMemUploadCounts();\n");
        script.append("  initSdErrorRatios();\n");
        script.append("  initSdErrorDetail();\n");
        script.append("});\n");
        script.append("function initTabs(sections) {\n");
        script.append("  const tabButtons = Array.from(document.querySelectorAll('.tab-link'));\n");
        script.append("  const panels = Array.from(document.querySelectorAll('.tab-panel'));\n");
        script.append("  const sectionIds = sections.map(section => section.id);\n");
        script.append("  const validIds = new Set(sectionIds);\n");
        script.append("  function activateTab(id, options) {\n");
        script.append("    if (!id || !validIds.has(id)) { return; }\n");
        script.append("    const opts = Object.assign({ updateHash: true, focus: false }, options || {});\n");
        script.append("    tabButtons.forEach(button => {\n");
        script.append("      const isActive = button.dataset.tab === id;\n");
        script.append("      button.classList.toggle('active', isActive);\n");
        script.append("      button.setAttribute('aria-selected', String(isActive));\n");
        script.append("      button.setAttribute('tabindex', isActive ? '0' : '-1');\n");
        script.append("    });\n");
        script.append("    panels.forEach(panel => {\n");
        script.append("      const isActive = panel.dataset.tabPanel === id;\n");
        script.append("      panel.classList.toggle('active', isActive);\n");
        script.append("      panel.setAttribute('aria-hidden', String(!isActive));\n");
        script.append("      panel.setAttribute('tabindex', isActive ? '0' : '-1');\n");
        script.append("    });\n");
        script.append("    if (opts.updateHash && typeof history.replaceState === 'function') {\n");
        script.append("      history.replaceState(null, '', '#' + id);\n");
        script.append("    }\n");
        script.append("    if (opts.focus) {\n");
        script.append("      const activeButton = tabButtons.find(button => button.dataset.tab === id);\n");
        script.append("      if (activeButton) { activeButton.focus(); }\n");
        script.append("    }\n");
        script.append("    setTimeout(() => {\n");
        script.append("      const panel = document.querySelector(`[data-tab-panel='${id}']`);\n");
        script.append("      if (panel) {\n");
        script.append("        panel.querySelectorAll('table').forEach(table => {\n");
        script.append("          const instance = dataTables[table.id];\n");
        script.append("          if (instance && typeof instance.columns === 'function') { instance.columns.adjust(); }\n");
        script.append("        });\n");
        script.append("        panel.querySelectorAll('canvas').forEach(canvas => {\n");
        script.append("          if (canvas.__chart && typeof canvas.__chart.resize === 'function') { canvas.__chart.resize(); }\n");
        script.append("        });\n");
        script.append("      }\n");
        script.append("    }, 150);\n");
        script.append("  }\n");
        script.append("  tabButtons.forEach(button => {\n");
        script.append("    button.addEventListener('click', () => activateTab(button.dataset.tab));\n");
        script.append("    button.addEventListener('keydown', event => {\n");
        script.append("      if (event.key === 'ArrowRight' || event.key === 'ArrowLeft') {\n");
        script.append("        event.preventDefault();\n");
        script.append("        const currentIndex = tabButtons.indexOf(button);\n");
        script.append("        const direction = event.key === 'ArrowRight' ? 1 : -1;\n");
        script.append("        const nextIndex = (currentIndex + direction + tabButtons.length) % tabButtons.length;\n");
        script.append("        const nextButton = tabButtons[nextIndex];\n");
        script.append("        if (nextButton) {\n");
        script.append("          activateTab(nextButton.dataset.tab, { focus: true });\n");
        script.append("        }\n");
        script.append("      }\n");
        script.append("    });\n");
        script.append("  });\n");
        script.append("  const hash = window.location.hash.replace('#', '');\n");
        script.append("  const initialId = validIds.has(hash) ? hash : sectionIds[0];\n");
        script.append("  if (initialId) { activateTab(initialId, { updateHash: false }); }\n");
        script.append("  window.addEventListener('hashchange', () => {\n");
        script.append("    const targetId = window.location.hash.replace('#', '');\n");
        script.append("    activateTab(targetId, { updateHash: false });\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("function initDataTable(tableId, columns, data, options) {\n");
        script.append("  const tableElement = document.getElementById(tableId);\n");
        script.append("  if (!tableElement) { return null; }\n");
        script.append("  const tableOptions = options || {};\n");
        script.append("  const exportTitle = tableOptions.exportTitle || tableId;\n");
        script.append("  const config = Object.assign({\n");
        script.append("    data: data,\n");
        script.append("    columns: columns.map(name => ({ title: name, data: name })),\n");
        script.append("    paging: true,\n");
        script.append("    searching: true,\n");
        script.append("    ordering: true,\n");
        script.append("    pageLength: 10,\n");
        script.append("    dom: 'Bfrtip',\n");
        script.append("    buttons: [\n");
        script.append("      { extend: 'excelHtml5', title: exportTitle },\n");
        script.append("      { extend: 'pdfHtml5', title: exportTitle, orientation: 'landscape', pageSize: 'A4' }\n");
        script.append("    ]\n");
        script.append("  }, tableOptions);\n");
        script.append("  const table = $(tableElement).DataTable(config);\n");
        script.append("  dataTables[tableId] = table;\n");
        script.append("  return table;\n");
        script.append("}\n");
        script.append("function safeNumber(value) {\n");
        script.append("  if (value === undefined || value === null) { return 0; }\n");
        script.append("  const normalised = String(value).replace(/,/g, '').trim();\n");
        script.append("  const num = Number(normalised);\n");
        script.append("  return Number.isFinite(num) ? num : 0;\n");
        script.append("}\n");
        script.append("function createChart(canvasId, config) {\n");
        script.append("  const canvas = document.getElementById(canvasId);\n");
        script.append("  if (!canvas) { return null; }\n");
        script.append("  const chart = new Chart(canvas, config);\n");
        script.append("  canvas.__chart = chart;\n");
        script.append("  return chart;\n");
        script.append("}\n");
        script.append("function initAbnormalIds() {\n");
        script.append("  if (!abnormalIdsData.length) {\n");
        script.append("    document.getElementById('abnormalIdsEmpty').classList.remove('hidden');\n");
        script.append("    document.getElementById('abnormalIdsChartContainer').classList.add('hidden');\n");
        script.append("    return;\n");
        script.append("  }\n");
        script.append("  initDataTable('abnormalIdsTable', abnormalIdsColumns, abnormalIdsData, {\n");
        script.append("    exportTitle: 'Abnormal Member IDs',\n");
        script.append("    pageLength: 15\n");
        script.append("  });\n");
        script.append("  const labels = abnormalIdsData.map(row => row['InsuranceCompanyName']);\n");
        script.append("  const data = abnormalIdsData.map(row => safeNumber(row['PolicyCount']));\n");
        script.append("  createChart('abnormalIdsChart', {\n");
        script.append("    type: 'bar',\n");
        script.append("    data: { labels: labels, datasets: [{\n");
        script.append("      label: 'Policies with invalid IDs',\n");
        script.append("      data: data,\n");
        script.append("      backgroundColor: '#38bdf8'\n");
        script.append("    }]},\n");
        script.append("    options: { responsive: true, plugins: { legend: { display: false }, tooltip: { mode: 'index' } }, scales: { y: { beginAtZero: true, title: { display: true, text: 'Policies' } } } }\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("function initIcpApiStats() {\n");
        script.append("  if (!icpApiData.length) {\n");
        script.append("    document.getElementById('icpApiEmpty').classList.remove('hidden');\n");
        script.append("    document.getElementById('icpApiChartContainer').classList.add('hidden');\n");
        script.append("    return;\n");
        script.append("  }\n");
        script.append("  initDataTable('icpApiTable', icpApiColumns, icpApiData, {\n");
        script.append("    exportTitle: 'ICP Service Success vs Failure',\n");
        script.append("    paging: false\n");
        script.append("  });\n");
        script.append("  const labels = icpApiData.map(row => row['ServiceName']);\n");
        script.append("  const success = icpApiData.map(row => safeNumber(row['SuccessCount']));\n");
        script.append("  const failure = icpApiData.map(row => safeNumber(row['FailureCount']));\n");
        script.append("  createChart('icpApiChart', {\n");
        script.append("    type: 'bar',\n");
        script.append("    data: { labels: labels, datasets: [\n");
        script.append("      { label: 'Success', data: success, backgroundColor: '#22c55e' },\n");
        script.append("      { label: 'Failure', data: failure, backgroundColor: '#ef4444' }\n");
        script.append("    ]},\n");
        script.append("    options: { responsive: true, plugins: { tooltip: { mode: 'index', intersect: false } }, scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } } }\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("function initIcpErrorDetails() {\n");
        script.append("  if (!icpErrorData.length) {\n");
        script.append("    document.getElementById('icpErrorsEmpty').classList.remove('hidden');\n");
        script.append("    return;\n");
        script.append("  }\n");
        script.append("  initDataTable('icpErrorsTable', icpErrorColumns, icpErrorData, {\n");
        script.append("    exportTitle: 'ICP Failure Reasons',\n");
        script.append("    pageLength: 15\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("function initMemUploadCounts() {\n");
        script.append("  if (!memUploadData.length) {\n");
        script.append("    document.getElementById('memUploadEmpty').classList.remove('hidden');\n");
        script.append("    document.getElementById('memUploadChartContainer').classList.add('hidden');\n");
        script.append("    return;\n");
        script.append("  }\n");
        script.append("  initDataTable('memUploadTable', memUploadColumns, memUploadData, {\n");
        script.append("    exportTitle: 'Policy Upload Channels',\n");
        script.append("    createdRow: function (row, data) {\n");
        script.append("      if (safeNumber(data['Manual_Upload']) > safeNumber(data['API_Upload'])) {\n");
        script.append("        row.classList.add('highlight-manual');\n");
        script.append("      }\n");
        script.append("    }\n");
        script.append("  });\n");
        script.append("  const labels = memUploadData.map(row => row['InsuranceCompanyName']);\n");
        script.append("  const apiCounts = memUploadData.map(row => safeNumber(row['API_Upload']));\n");
        script.append("  const manualCounts = memUploadData.map(row => safeNumber(row['Manual_Upload']));\n");
        script.append("  createChart('memUploadChart', {\n");
        script.append("    type: 'bar',\n");
        script.append("    data: { labels: labels, datasets: [\n");
        script.append("      { label: 'API Upload', data: apiCounts, backgroundColor: '#3b82f6' },\n");
        script.append("      { label: 'Manual Upload', data: manualCounts, backgroundColor: '#f97316' }\n");
        script.append("    ]},\n");
        script.append("    options: { responsive: true, plugins: { tooltip: { mode: 'index', intersect: false } }, scales: { x: { stacked: false, title: { display: true, text: 'Insurance Company' } }, y: { beginAtZero: true, title: { display: true, text: 'Policies' } } } }\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("function initSdErrorRatios() {\n");
        script.append("  if (!sdErrorData.length) {\n");
        script.append("    document.getElementById('sdErrorEmpty').classList.remove('hidden');\n");
        script.append("    document.getElementById('sdErrorChartContainer').classList.add('hidden');\n");
        script.append("    return;\n");
        script.append("  }\n");
        script.append("  initDataTable('sdErrorTable', sdErrorColumns, sdErrorData, {\n");
        script.append("    exportTitle: 'Service Failure Ratios',\n");
        script.append("    pageLength: 15\n");
        script.append("  });\n");
        script.append("  const labels = sdErrorData.map(row => row['InsuranceCompanyName']);\n");
        script.append("  const ratios = sdErrorData.map(row => safeNumber(row['API_Failure_Ratio']));\n");
        script.append("  createChart('sdErrorChart', {\n");
        script.append("    type: 'line',\n");
        script.append("    data: { labels: labels, datasets: [{\n");
        script.append("      label: 'Failure Ratio',\n");
        script.append("      data: ratios,\n");
        script.append("      borderColor: '#f97316',\n");
        script.append("      backgroundColor: 'rgba(249, 115, 22, 0.2)',\n");
        script.append("      fill: true,\n");
        script.append("      tension: 0.3\n");
        script.append("    }]},\n");
        script.append("    options: { responsive: true, plugins: { legend: { position: 'top' } }, scales: { y: { beginAtZero: true } } }\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("function initSdErrorDetail() {\n");
        script.append("  if (!sdErrorDetailData.length) {\n");
        script.append("    document.getElementById('sdErrorDetailEmpty').classList.remove('hidden');\n");
        script.append("    document.getElementById('sdErrorDetailChartContainer').classList.add('hidden');\n");
        script.append("    return;\n");
        script.append("  }\n");
        script.append("  initDataTable('sdErrorDetailTable', sdErrorDetailColumns, sdErrorDetailData, {\n");
        script.append("    exportTitle: 'Error Details by Insurance Company',\n");
        script.append("    pageLength: 20\n");
        script.append("  });\n");
        script.append("  const aggregation = {};\n");
        script.append("  sdErrorDetailData.forEach(row => {\n");
        script.append("    const key = row['Error_Description'] || row['ServiceName'] || 'Unknown';\n");
        script.append("    aggregation[key] = (aggregation[key] || 0) + safeNumber(row['Error_Count']);\n");
        script.append("  });\n");
        script.append("  const topEntries = Object.entries(aggregation).sort((a, b) => b[1] - a[1]).slice(0, 10);\n");
        script.append("  const labels = topEntries.map(entry => entry[0]);\n");
        script.append("  const values = topEntries.map(entry => entry[1]);\n");
        script.append("  createChart('sdErrorDetailChart', {\n");
        script.append("    type: 'bar',\n");
        script.append("    data: { labels: labels, datasets: [{\n");
        script.append("      label: 'Error Count',\n");
        script.append("      data: values,\n");
        script.append("      backgroundColor: '#a855f7'\n");
        script.append("    }]},\n");
        script.append("    options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } }\n");
        script.append("  });\n");
        script.append("}\n");
        script.append("\n");
        return script.toString();
    }
}
