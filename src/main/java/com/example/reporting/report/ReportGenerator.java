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
                .append("<div class=\"brand-identity\">")
                .append("<div class=\"brand-logo\" role=\"img\" aria-label=\"SlashData logo\">")
                .append("<svg viewBox=\"0 0 64 64\" xmlns=\"http://www.w3.org/2000/svg\" focusable=\"false\" aria-hidden=\"true\">")
                .append("<path d=\"M45.6 4H60L18.4 60H4z\" fill=\"#7fb3ff\" />")
                .append("<path d=\"M38.6 4H52L10.4 60H-3.6z\" fill=\"rgba(255,255,255,0.18)\" />")
                .append("</svg>")
                .append("</div>")
                .append("<div class=\"brand-meta\">")
                .append("<span class=\"company-name\">SlashData</span>")
                .append("<span class=\"product-name\">Product: Rabet</span>")
                .append("</div>")
                .append("</div>")
                .append("<span class=\"page-badge\">Weekly Insight</span>")
                .append("<h1>SD Interactive Reporting Dashboard</h1>")
                .append("<p class=\"subtitle\">Operational quality metrics for the Rabet service desk platform.</p>")
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

        html.append(renderSectionIconBar());

        html.append("<main class=\"content\">")
                .append(renderKeyMetrics(summary));

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
                .append("<script type=\"text/javascript\">\n")
                .append("//<![CDATA[\n")
                .append("const abnormalIdsColumns = ").append(toJsonArray(abnormalColumns)).append(";\n")
                .append("const abnormalIdsData = ").append(abnormalJson).append(";\n")
                .append("const icpApiColumns = ").append(toJsonArray(icpApiColumns)).append(";\n")
                .append("const icpApiData = ").append(icpApiJson).append(";\n")
                .append("const icpErrorColumns = ").append(toJsonArray(icpErrorColumns)).append(";\n")
                .append("const icpErrorData = ").append(icpErrorJson).append(";\n")
                .append("const memUploadColumns = ").append(toJsonArray(memUploadColumns)).append(";\n")
                .append("const memUploadData = ").append(memUploadJson).append(";\n")
                .append("const sdErrorColumns = ").append(toJsonArray(sdErrorColumns)).append(";\n")
                .append("const sdErrorData = ").append(sdErrorJson).append(";\n")
                .append("const sdErrorDetailColumns = ").append(toJsonArray(sdErrorDetailColumns)).append(";\n")
                .append("const sdErrorDetailData = ").append(sdErrorDetailJson).append(";\n")
                .append(getScript())
                .append("//]]>\n")
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

    private String renderDataSection(String sectionId,
                                     String title,
                                     String description,
                                     String chartContainerId,
                                     String chartId,
                                     String tableId,
                                     String emptyMessageId) {
        String classes = "card data-section data-section--with-chart";
        return new StringBuilder()
                .append('<').append("section class=\"").append(classes).append("\" id=\"").append(sectionId)
                .append("\">")
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
        String classes = "card data-section data-section--table-only";
        return new StringBuilder()
                .append('<').append("section class=\"").append(classes).append("\" id=\"").append(sectionId)
                .append("\">")
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

    private String renderSectionIconBar() {
        StringBuilder builder = new StringBuilder()
                .append("<nav class=\"section-icon-bar\" aria-label=\"Report section shortcuts\">");
        for (String[] section : REPORT_SECTIONS) {
            String id = section[0];
            String label = section[1];
            builder.append(String.format(
                    Locale.ROOT,
                    "<a class=\"section-icon-link\" href=\"#%s\" aria-label=\"%s\"><span class=\"section-icon-symbol\" aria-hidden=\"true\">%s</span><span class=\"section-icon-label\">%s</span></a>",
                    escapeHtml(id),
                    escapeHtml(label),
                    escapeHtml(getSectionIconSymbol(id)),
                    escapeHtml(label)
            ));
        }
        builder.append("</nav>");
        return builder.toString();
    }

    private String getSectionIconSymbol(String sectionId) {
        return switch (sectionId) {
            case "abnormal-ids" -> "ID";
            case "icp-api-stats" -> "API";
            case "icp-error-details" -> "ERR";
            case "mem-upload-counts" -> "UP";
            case "sd-error-ratio" -> "SFR";
            case "sd-error-details-ic" -> "IC";
            default -> "DATA";
        };
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
        StringBuilder styles = new StringBuilder();
        styles.append(":root { --page-bg: #f4f6fb; --surface: #ffffff; --surface-muted: #f8fafc; --surface-raised: #eef2ff; --ink-900: #0f172a; --ink-700: #1f2937; --ink-500: #475569; --ink-400: #64748b; --accent-50: #eff6ff; --accent-100: #dbeafe; --accent-200: #bfdbfe; --accent-400: #60a5fa; --accent-500: #3b82f6; --accent-600: #2563eb; --accent-700: #1d4ed8; --border-subtle: #e2e8f0; --border-strong: #cbd5f5; --shadow-card: 0 22px 40px -28px rgba(15, 23, 42, 0.45); --warning-100: #fef3c7; --warning-600: #d97706; --danger-100: #fee2e2; --danger-600: #dc2626; }");
        styles.append("*, *::before, *::after { box-sizing: border-box; }");
        styles.append("body { margin: 0; font-family: 'Segoe UI', Arial, sans-serif; background: var(--page-bg); color: var(--ink-900); line-height: 1.55; }");
        styles.append("a { color: var(--accent-600); text-decoration: none; }");
        styles.append("a:hover { text-decoration: underline; }");
        styles.append(".page-shell { max-width: 1240px; margin: 0 auto; padding: 40px 24px 64px; }");
        styles.append(".page-header { display: grid; gap: 24px; background: var(--ink-900); color: var(--surface); padding: 36px 32px; border-radius: 20px; margin-bottom: 48px; border: 1px solid rgba(15, 23, 42, 0.3); position: relative; overflow: hidden; }");
        styles.append(".page-header::after { content: ''; position: absolute; inset: 0; background: linear-gradient(135deg, rgba(59, 130, 246, 0.35), transparent 58%); pointer-events: none; }");
        styles.append(".page-header > * { position: relative; z-index: 1; }");
        styles.append(".branding { margin-bottom: 8px; }");
        styles.append(".brand-identity { display: flex; align-items: center; gap: 16px; margin-bottom: 18px; }");
        styles.append(".brand-logo { width: 64px; height: 64px; border-radius: 20px; background: rgba(148, 163, 184, 0.18); display: flex; align-items: center; justify-content: center; box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.2); }");
        styles.append(".brand-logo svg { width: 40px; height: 40px; display: block; }");
        styles.append(".brand-meta { display: flex; flex-direction: column; gap: 4px; }");
        styles.append(".company-name { font-size: 18px; font-weight: 600; letter-spacing: 0.18em; text-transform: uppercase; color: var(--surface); }");
        styles.append(".product-name { font-size: 13px; letter-spacing: 0.26em; text-transform: uppercase; color: var(--accent-200); }");
        styles.append(".page-badge { display: inline-flex; align-items: center; gap: 6px; background: rgba(59, 130, 246, 0.18); color: var(--surface); border-radius: 999px; padding: 6px 14px; font-size: 12px; letter-spacing: 0.14em; text-transform: uppercase; }");
        styles.append(".page-badge::before { content: ''; width: 8px; height: 8px; border-radius: 999px; background: var(--accent-400); }");
        styles.append(".branding h1 { margin: 0; font-size: 32px; font-weight: 600; letter-spacing: -0.02em; }");
        styles.append(".subtitle { margin: 12px 0 0; font-size: 16px; color: var(--accent-100); max-width: 520px; }");
        styles.append(".meta-grid { display: grid; gap: 16px; margin: 4px 0 0; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); }");
        styles.append(".meta-item { background: rgba(15, 23, 42, 0.55); border-radius: 16px; padding: 16px 20px; border: 1px solid rgba(148, 163, 184, 0.38); backdrop-filter: blur(4px); }");
        styles.append(".meta-label { display: block; font-size: 11px; text-transform: uppercase; letter-spacing: 0.16em; color: var(--accent-200); margin-bottom: 6px; }");
        styles.append(".meta-value { font-size: 20px; font-weight: 600; color: var(--surface); word-break: break-word; }");
        styles.append(".callout { background: rgba(15, 23, 42, 0.65); border-radius: 16px; padding: 16px 20px; font-size: 15px; color: var(--accent-100); border: 1px solid rgba(148, 163, 184, 0.4); }");
        styles.append(".content { margin-top: 32px; }");
        styles.append(".content > section { margin-bottom: 32px; }");
        styles.append(".section-icon-bar { display: flex; flex-wrap: wrap; gap: 18px; margin: 36px 0 0; padding: 0; list-style: none; }");
        styles.append(".section-icon-link { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 8px; text-decoration: none; color: var(--ink-700); min-width: 84px; }");
        styles.append(".section-icon-symbol { display: inline-flex; align-items: center; justify-content: center; width: 60px; height: 60px; border-radius: 18px; font-weight: 700; letter-spacing: 0.08em; background: linear-gradient(135deg, var(--accent-600), var(--accent-400)); color: var(--surface); box-shadow: 0 12px 24px -20px rgba(37, 99, 235, 0.9); transition: transform 0.15s ease, box-shadow 0.15s ease; font-size: 15px; text-transform: uppercase; }");
        styles.append(".section-icon-link:hover .section-icon-symbol { transform: translateY(-2px); box-shadow: 0 16px 28px -20px rgba(37, 99, 235, 0.95); }");
        styles.append(".section-icon-link:focus-visible .section-icon-symbol { outline: 2px solid var(--accent-400); outline-offset: 4px; }");
        styles.append(".section-icon-label { font-size: 13px; color: var(--ink-500); text-align: center; max-width: 110px; line-height: 1.4; }");
        styles.append(".card { background: var(--surface); border-radius: 20px; padding: 28px; border: 1px solid var(--border-subtle); box-shadow: var(--shadow-card); }");
        styles.append(".card h2 { margin: 0; font-size: 22px; font-weight: 600; color: var(--ink-900); }");
        styles.append(".metrics-grid { display: grid; gap: 18px; margin-top: 20px; grid-template-columns: minmax(0, 1fr); }");
        styles.append(".metric { padding: 18px 20px; border-radius: 16px; background: var(--accent-50); border: 1px solid var(--border-strong); display: grid; gap: 6px; }");
        styles.append(".metric-value { font-size: 26px; font-weight: 600; color: var(--ink-900); }");
        styles.append(".metric-label { font-size: 11px; letter-spacing: 0.16em; text-transform: uppercase; color: var(--ink-500); }");
        styles.append(".metric-caption { font-size: 14px; color: var(--ink-400); }");
        styles.append(".missing { border-left: 4px solid var(--warning-600); padding-left: 18px; background: var(--warning-100); border-radius: 16px; }");
        styles.append(".missing h2 { color: var(--warning-600); }");
        styles.append(".missing p { margin-top: 12px; color: var(--ink-700); }");
        styles.append(".missing ul { margin: 16px 0 0; padding-left: 20px; color: var(--ink-700); }");
        styles.append(".data-section { grid-template-columns: minmax(0, 1fr); gap: 24px; align-content: start; }");
        styles.append(".data-section--with-chart .empty { text-align: center; }");
        styles.append(".section-header { display: grid; gap: 8px; margin-bottom: 14px; }");
        styles.append(".section-header h2 { margin: 0; font-size: 20px; }");
        styles.append(".description { margin: 0; color: var(--ink-400); line-height: 1.6; }");
        styles.append(".chart-card { min-height: 300px; padding: 18px; background: var(--surface-muted); border-radius: 16px; border: 1px solid var(--border-strong); display: flex; align-items: center; justify-content: center; }");
        styles.append(".chart-card canvas { width: 100% !important; height: 100% !important; }");
        styles.append(".table-card { border-radius: 16px; border: 1px solid var(--border-subtle); background: var(--surface); padding: 0; overflow-x: auto; }");
        styles.append(".table-card table { width: 100%; }");
        styles.append("table.dataTable { width: 100% !important; border-collapse: collapse; font-size: 15px; color: var(--ink-700); }");
        styles.append("table.dataTable th, table.dataTable td { border: 1px solid var(--border-strong); padding: 10px 12px; text-align: left; }");
        styles.append("table.dataTable thead th { background: var(--accent-100); color: var(--ink-900); font-weight: 600; }");
        styles.append("table.dataTable tbody tr:nth-child(even) { background: var(--surface-muted); }");
        styles.append(".dataTables_wrapper { width: 100%; padding: 20px; }");
        styles.append(".dataTables_wrapper .dataTables_filter { text-align: left; color: var(--ink-500); }");
        styles.append(".dataTables_wrapper .dataTables_filter input { margin-left: 0.5em; padding: 6px 12px; border-radius: 999px; border: 1px solid var(--border-strong); background: var(--surface); }");
        styles.append(".dataTables_wrapper .dataTables_info { color: var(--ink-500); }");
        styles.append(".dataTables_wrapper .dataTables_paginate { display: flex; gap: 8px; justify-content: flex-end; flex-wrap: wrap; padding-top: 12px; }");
        styles.append(".dataTables_wrapper .dataTables_paginate .paginate_button { border-radius: 999px; border: 1px solid var(--border-strong); background: var(--surface); color: var(--ink-700) !important; padding: 6px 14px; margin: 0; transition: all 0.15s ease; }");
        styles.append(".dataTables_wrapper .dataTables_paginate .paginate_button:hover { border-color: var(--accent-400); color: var(--ink-900) !important; background: var(--accent-100); }");
        styles.append(".dataTables_wrapper .dataTables_paginate .paginate_button.current { background: var(--accent-600) !important; color: var(--surface) !important; border-color: transparent; }");
        styles.append(".dt-buttons { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 14px; }");
        styles.append(".dt-button { background: var(--accent-600); border: none; color: var(--surface); border-radius: 999px; padding: 8px 18px; font-weight: 600; cursor: pointer; transition: background 0.15s ease; box-shadow: 0 10px 24px -18px rgba(37, 99, 235, 0.9); }");
        styles.append(".dt-button:hover { background: var(--accent-700); }");
        styles.append(".dt-button:focus-visible { outline: 2px solid var(--accent-400); outline-offset: 3px; }");
        styles.append(".empty { color: var(--ink-400); font-style: italic; margin: 0; }");
        styles.append(".hidden { display: none; }");
        styles.append(".highlight-manual { background-color: var(--warning-100) !important; }");
        styles.append("footer { margin-top: 48px; text-align: center; color: var(--ink-400); font-size: 14px; }");
        styles.append("@media (max-width: 599px) { .page-header { padding: 30px 24px; } .page-shell { padding: 32px 18px 56px; } .card { padding: 22px; } .dataTables_wrapper { padding: 16px; } .section-icon-bar { justify-content: center; } .section-icon-link { min-width: 72px; } }");
        styles.append("@media (min-width: 600px) { .branding h1 { font-size: 34px; } }");
        styles.append("@media (min-width: 768px) { .page-shell { padding: 48px 40px 64px; } .card { padding: 30px; } .metrics-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } .meta-grid { grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); } }");
        styles.append("@media (min-width: 1100px) { .metrics-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); } }");
        styles.append("@media (min-width: 1024px) { .page-header { grid-template-columns: minmax(0, 2fr) minmax(0, 1fr); grid-template-areas: 'branding meta' 'callout meta'; align-items: start; } .page-header .branding { grid-area: branding; margin-bottom: 0; } .page-header .meta-grid { grid-area: meta; margin: 0; align-self: stretch; } .page-header .callout { grid-area: callout; margin-top: 0; } }");
        styles.append("@media (min-width: 1024px) { .data-section--with-chart { grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); grid-template-areas: 'header header' 'chart table'; align-items: stretch; } .data-section--with-chart .section-header { grid-area: header; } .data-section--with-chart .chart-card { grid-area: chart; } .data-section--with-chart .table-card { grid-area: table; } .data-section--with-chart .empty { grid-area: chart; } }");
        styles.append("@media (max-width: 480px) { .dataTables_wrapper .dataTables_filter input { width: 100%; margin-top: 8px; } .dataTables_wrapper .dataTables_filter label { display: flex; flex-direction: column; align-items: stretch; gap: 6px; } }");
        return styles.toString();
    }

    private String getScript() {
        StringBuilder script = new StringBuilder();
        script.append("const dataTables = {};\n");
        script.append("document.addEventListener('DOMContentLoaded', function () {\n");
        script.append("  initAbnormalIds();\n");
        script.append("  initIcpApiStats();\n");
        script.append("  initIcpErrorDetails();\n");
        script.append("  initMemUploadCounts();\n");
        script.append("  initSdErrorRatios();\n");
        script.append("  initSdErrorDetail();\n");
        script.append("});\n");
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
        script.append("    autoWidth: false,\n");
        script.append("    scrollX: true,\n");
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
