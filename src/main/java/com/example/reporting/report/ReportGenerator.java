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
                .append("<body>");

        html.append("<header>")
                .append("<div class=\"title-block\">")
                .append("<h1>SD Interactive Reporting Dashboard</h1>")
                .append(String.format("<p class=\"meta\"><strong>Report Date:</strong> %s</p>",
                        DATE_FORMATTER.format(summary.getReportDate())))
                .append(String.format("<p class=\"meta\"><strong>Generated At:</strong> %s</p>",
                        escapeHtml(summary.getGeneratedAt().toString())))
                .append(String.format("<p class=\"meta\"><strong>Total Rows Processed:</strong> %,d</p>",
                        summary.getTotalRowCount()))
                .append("</div>")
                .append("<div class=\"export-reminder\">")
                .append("<p>Use the export buttons on each table to download Excel or PDF extracts.</p>")
                .append("</div>")
                .append("</header>");

        if (!summary.getMissingSheets().isEmpty()) {
            html.append("<section class=\"missing\">")
                    .append("<h2>Missing Worksheets</h2>")
                    .append("<ul>");
            summary.getMissingSheets().forEach(sheet ->
                    html.append(String.format("<li>%s</li>", escapeHtml(sheet))));
            html.append("</ul></section>");
        }

        html.append("<section class=\"overview\">")
                .append("<article class=\"card\">")
                .append("<h2>Report Highlights</h2>")
                .append("<p>Interactive charts and smart tables below help you investigate policy uploads, API performance, and error trends. "
                        + "Hover over chart elements for tooltips, filter tables using the search boxes, and export data for offline analysis.</p>")
                .append("</article>")
                .append("</section>");

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

        html.append("<footer><p class=\"meta\">Report generated automatically from the latest workbook.</p></footer>");

        html.append("<script src=\"https://code.jquery.com/jquery-3.7.1.min.js\"></script>")
                .append("<script src=\"https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js\"></script>")
                .append("<script src=\"https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js\"></script>")
                .append("<script src=\"https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js\"></script>")
                .append("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js\"></script>")
                .append("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js\"></script>")
                .append("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js\"></script>")
                .append("<script src=\"https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js\"></script>")
                .append("<script>")
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

    private String renderDataSection(String sectionId,
                                     String title,
                                     String description,
                                     String chartContainerId,
                                     String chartId,
                                     String tableId,
                                     String emptyMessageId) {
        return new StringBuilder()
                .append('<').append("section class=\"card\" id=\"").append(sectionId).append("\">")
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
        return new StringBuilder()
                .append('<').append("section class=\"card\" id=\"").append(sectionId).append("\">")
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
        return "body { font-family: 'Segoe UI', Arial, sans-serif; margin: 1.5rem; background-color: #f8fafc; color: #0f172a; }"
                + "header { display: flex; flex-direction: column; gap: 0.8rem; padding-bottom: 1rem; border-bottom: 2px solid #e2e8f0; margin-bottom: 2rem; }"
                + ".title-block h1 { margin: 0; font-size: 2rem; }"
                + ".meta { margin: 0.2rem 0; color: #475569; }"
                + ".export-reminder { background: #e0f2fe; border-left: 4px solid #0284c7; padding: 0.75rem 1rem; border-radius: 4px; color: #0c4a6e; font-size: 0.95rem; }"
                + "section { margin-bottom: 2rem; }"
                + ".card { background: #ffffff; border-radius: 12px; padding: 1.5rem; box-shadow: 0 15px 25px -20px rgba(15, 23, 42, 0.45); }"
                + ".section-header { margin-bottom: 1rem; }"
                + ".section-header h2 { margin: 0 0 0.4rem; font-size: 1.4rem; }"
                + ".description { margin: 0; color: #475569; }"
                + ".chart-card { position: relative; min-height: 320px; padding: 1rem; background: #f1f5f9; border-radius: 10px; margin-bottom: 1rem; }"
                + ".chart-card canvas { width: 100% !important; height: 100% !important; }"
                + ".table-card { overflow-x: auto; }"
                + "table.dataTable { width: 100% !important; border-collapse: collapse; }"
                + "table.dataTable thead th { background: #e2e8f0; color: #0f172a; }"
                + "table.dataTable tbody tr:nth-child(even) { background: #f8fafc; }"
                + ".missing ul { list-style: disc; padding-left: 1.5rem; }"
                + ".empty { color: #64748b; font-style: italic; margin: 0; }"
                + ".hidden { display: none; }"
                + ".highlight-manual { background-color: #fef3c7 !important; }"
                + "footer { margin-top: 3rem; text-align: center; color: #64748b; font-size: 0.85rem; }"
                + "@media (min-width: 1024px) { header { flex-direction: row; justify-content: space-between; align-items: flex-end; } }";
    }

    private String getScript() {
        return "document.addEventListener('DOMContentLoaded', function () {"
                + "initAbnormalIds();"
                + "initIcpApiStats();"
                + "initIcpErrorDetails();"
                + "initMemUploadCounts();"
                + "initSdErrorRatios();"
                + "initSdErrorDetail();"
                + "});"
                + "function initDataTable(tableId, columns, data, options) {"
                + "const tableElement = document.getElementById(tableId);"
                + "if (!tableElement) { return null; }"
                + "const config = Object.assign({"
                + "data: data,"
                + "columns: columns.map(name => ({ title: name, data: name })),"
                + "paging: true,"
                + "searching: true,"
                + "ordering: true,"
                + "pageLength: 10,"
                + "dom: 'Bfrtip',"
                + "buttons: ["
                + "{ extend: 'excelHtml5', title: options.exportTitle || tableId },"
                + "{ extend: 'pdfHtml5', title: options.exportTitle || tableId, orientation: 'landscape', pageSize: 'A4' }"
                + " ]"
                + "}, options || {});"
                + "return $(tableElement).DataTable(config);"
                + "}"
                + "function safeNumber(value) {"
                + "if (value === undefined || value === null) { return 0; }"
                + "const normalised = String(value).replace(/,/g, '').trim();"
                + "const num = Number(normalised);"
                + "return Number.isFinite(num) ? num : 0;"
                + "}"
                + "function initAbnormalIds() {"
                + "if (!abnormalIdsData.length) {"
                + "document.getElementById('abnormalIdsEmpty').classList.remove('hidden');"
                + "document.getElementById('abnormalIdsChartContainer').classList.add('hidden');"
                + "return;"
                + "}"
                + "initDataTable('abnormalIdsTable', abnormalIdsColumns, abnormalIdsData, {"
                + "exportTitle: 'Abnormal Member IDs',"
                + "pageLength: 15"
                + "});"
                + "const labels = abnormalIdsData.map(row => row['InsuranceCompanyName']);"
                + "const data = abnormalIdsData.map(row => safeNumber(row['PolicyCount']));"
                + "new Chart(document.getElementById('abnormalIdsChart'), {"
                + "type: 'bar',"
                + "data: { labels: labels, datasets: [{"
                + "label: 'Policies with invalid IDs',"
                + "data: data,"
                + "backgroundColor: '#38bdf8'"
                + "}]},"
                + "options: { responsive: true, plugins: { legend: { display: false }, tooltip: { mode: 'index' } }, scales: { y: { beginAtZero: true, title: { display: true, text: 'Policies' } } } }"
                + "});"
                + "}"
                + "function initIcpApiStats() {"
                + "if (!icpApiData.length) {"
                + "document.getElementById('icpApiEmpty').classList.remove('hidden');"
                + "document.getElementById('icpApiChartContainer').classList.add('hidden');"
                + "return;"
                + "}"
                + "initDataTable('icpApiTable', icpApiColumns, icpApiData, {"
                + "exportTitle: 'ICP Service Success vs Failure',"
                + "paging: false"
                + "});"
                + "const labels = icpApiData.map(row => row['ServiceName']);"
                + "const success = icpApiData.map(row => safeNumber(row['SuccessCount']));"
                + "const failure = icpApiData.map(row => safeNumber(row['FailureCount']));"
                + "new Chart(document.getElementById('icpApiChart'), {"
                + "type: 'bar',"
                + "data: { labels: labels, datasets: ["
                + "{ label: 'Success', data: success, backgroundColor: '#22c55e' },"
                + "{ label: 'Failure', data: failure, backgroundColor: '#ef4444' }"
                + "]},"
                + "options: { responsive: true, plugins: { tooltip: { mode: 'index', intersect: false } }, scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } } }"
                + "});"
                + "}"
                + "function initIcpErrorDetails() {"
                + "if (!icpErrorData.length) {"
                + "document.getElementById('icpErrorsEmpty').classList.remove('hidden');"
                + "return;"
                + "}"
                + "initDataTable('icpErrorsTable', icpErrorColumns, icpErrorData, {"
                + "exportTitle: 'ICP Failure Reasons',"
                + "pageLength: 15"
                + "});"
                + "}"
                + "function initMemUploadCounts() {"
                + "if (!memUploadData.length) {"
                + "document.getElementById('memUploadEmpty').classList.remove('hidden');"
                + "document.getElementById('memUploadChartContainer').classList.add('hidden');"
                + "return;"
                + "}"
                + "initDataTable('memUploadTable', memUploadColumns, memUploadData, {"
                + "exportTitle: 'Policy Upload Channels',"
                + "createdRow: function (row, data) {"
                + "if (safeNumber(data['Manual_Upload']) > safeNumber(data['API_Upload'])) {"
                + "row.classList.add('highlight-manual');"
                + "}"
                + "}"
                + "});"
                + "const labels = memUploadData.map(row => row['InsuranceCompanyName']);"
                + "const apiCounts = memUploadData.map(row => safeNumber(row['API_Upload']));"
                + "const manualCounts = memUploadData.map(row => safeNumber(row['Manual_Upload']));"
                + "new Chart(document.getElementById('memUploadChart'), {"
                + "type: 'bar',"
                + "data: { labels: labels, datasets: ["
                + "{ label: 'API Upload', data: apiCounts, backgroundColor: '#3b82f6' },"
                + "{ label: 'Manual Upload', data: manualCounts, backgroundColor: '#f97316' }"
                + "]},"
                + "options: { responsive: true, plugins: { tooltip: { mode: 'index', intersect: false } }, scales: { x: { stacked: false, title: { display: true, text: 'Insurance Company' } }, y: { beginAtZero: true, title: { display: true, text: 'Policies' } } } }"
                + "});"
                + "}"
                + "function initSdErrorRatios() {"
                + "if (!sdErrorData.length) {"
                + "document.getElementById('sdErrorEmpty').classList.remove('hidden');"
                + "document.getElementById('sdErrorChartContainer').classList.add('hidden');"
                + "return;"
                + "}"
                + "initDataTable('sdErrorTable', sdErrorColumns, sdErrorData, {"
                + "exportTitle: 'Service Failure Ratios',"
                + "pageLength: 15"
                + "});"
                + "const labels = sdErrorData.map(row => row['InsuranceCompanyName']);"
                + "const ratios = sdErrorData.map(row => safeNumber(row['API_Failure_Ratio']));"
                + "new Chart(document.getElementById('sdErrorChart'), {"
                + "type: 'line',"
                + "data: { labels: labels, datasets: [{"
                + "label: 'Failure Ratio',"
                + "data: ratios,"
                + "borderColor: '#f97316',"
                + "backgroundColor: 'rgba(249, 115, 22, 0.2)',"
                + "fill: true,"
                + "tension: 0.3"
                + "}]},"
                + "options: { responsive: true, plugins: { legend: { position: 'top' } }, scales: { y: { beginAtZero: true } } }"
                + "});"
                + "}"
                + "function initSdErrorDetail() {"
                + "if (!sdErrorDetailData.length) {"
                + "document.getElementById('sdErrorDetailEmpty').classList.remove('hidden');"
                + "document.getElementById('sdErrorDetailChartContainer').classList.add('hidden');"
                + "return;"
                + "}"
                + "initDataTable('sdErrorDetailTable', sdErrorDetailColumns, sdErrorDetailData, {"
                + "exportTitle: 'Error Details by Insurance Company',"
                + "pageLength: 20"
                + "});"
                + "const aggregation = {};"
                + "sdErrorDetailData.forEach(row => {"
                + "const key = row['Error_Description'] || row['ServiceName'] || 'Unknown';"
                + "aggregation[key] = (aggregation[key] || 0) + safeNumber(row['Error_Count']);"
                + "});"
                + "const topEntries = Object.entries(aggregation).sort((a, b) => b[1] - a[1]).slice(0, 10);"
                + "const labels = topEntries.map(entry => entry[0]);"
                + "const values = topEntries.map(entry => entry[1]);"
                + "new Chart(document.getElementById('sdErrorDetailChart'), {"
                + "type: 'bar',"
                + "data: { labels: labels, datasets: [{"
                + "label: 'Error Count',"
                + "data: values,"
                + "backgroundColor: '#a855f7'"
                + "}]},"
                + "options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } }"
                + "});"
                + "}"
                + "";
    }
}
