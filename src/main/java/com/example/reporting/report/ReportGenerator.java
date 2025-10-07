package com.example.reporting.report;

import com.example.reporting.model.ReportSummary;
import com.example.reporting.model.SheetSummary;
import com.example.reporting.util.ChartRenderer;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.stream.Collectors;

import com.openhtmltopdf.pdfboxout.PdfRendererBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ReportGenerator {
    private static final Logger LOGGER = LoggerFactory.getLogger(ReportGenerator.class);
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("MMMM dd, yyyy", Locale.US);

    private final ChartRenderer chartRenderer;

    public ReportGenerator(ChartRenderer chartRenderer) {
        this.chartRenderer = chartRenderer;
    }

    public String buildHtmlReport(ReportSummary summary) {
        String chartData = chartRenderer.renderRowCountChart(summary.getSheetSummaries());

        StringBuilder html = new StringBuilder();
        html.append("<!DOCTYPE html>")
                .append("<html lang=\"en\">")
                .append("<head>")
                .append("<meta charset=\"UTF-8\" />")
                .append("<title>Weekly SD Report</title>")
                .append("<style>")
                .append(getStyles())
                .append("</style>")
                .append("</head>")
                .append("<body>");

        html.append("<header>")
                .append("<h1>SD Weekly Statistics Summary</h1>")
                .append(String.format("<p class=\"meta\">Report Date: %s</p>", DATE_FORMATTER.format(summary.getReportDate())))
                .append(String.format("<p class=\"meta\">Generated At: %s</p>", summary.getGeneratedAt()))
                .append(String.format("<p class=\"meta\">Total Rows Processed: %,d</p>", summary.getTotalRowCount()))
                .append("</header>");

        if (!summary.getMissingSheets().isEmpty()) {
            html.append("<section class=\"missing\">")
                    .append("<h2>Missing Worksheets</h2>")
                    .append("<ul>");
            summary.getMissingSheets().forEach(sheet ->
                    html.append(String.format("<li>%s</li>", escapeHtml(sheet))));
            html.append("</ul></section>");
        }

        if (!chartData.isBlank()) {
            html.append("<section class=\"chart\">")
                    .append("<h2>Row Volume by Worksheet</h2>")
                    .append(String.format("<img src=\"data:image/png;base64,%s\" alt=\"Row counts chart\" />", chartData))
                    .append("</section>");
        }

        html.append("<section class=\"details\">")
                .append("<h2>Worksheet Details</h2>");

        for (SheetSummary sheetSummary : summary.getSheetSummaries().values()) {
            html.append("<article class=\"sheet\">")
                    .append(String.format("<h3>%s</h3>", escapeHtml(sheetSummary.getSheetName())))
                    .append(String.format("<p><strong>Rows:</strong> %,d</p>", sheetSummary.getRowCount()))
                    .append(String.format("<p><strong>Columns:</strong> %,d</p>", sheetSummary.getColumnHeaders().size()));

            if (!sheetSummary.getSampleRows().isEmpty()) {
                html.append(renderTable(sheetSummary));
            } else {
                html.append("<p class=\"empty\">No sample data available for this worksheet.</p>");
            }

            html.append("</article>");
        }

        html.append("</section>")
                .append("</body></html>");

        return html.toString();
    }

    private String renderTable(SheetSummary summary) {
        String headers = summary.getColumnHeaders().stream()
                .map(this::escapeHtml)
                .map(value -> String.format("<th>%s</th>", value))
                .collect(Collectors.joining());

        String rows = summary.getSampleRows().stream()
                .map(row -> row.stream()
                        .map(this::escapeHtml)
                        .map(value -> String.format("<td>%s</td>", value))
                        .collect(Collectors.joining()))
                .map(cells -> String.format("<tr>%s</tr>", cells))
                .collect(Collectors.joining());

        return String.format(Locale.US,
                "<table><thead><tr>%s</tr></thead><tbody>%s</tbody></table>",
                headers,
                rows);
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

    private String getStyles() {
        return "body { font-family: 'Segoe UI', Arial, sans-serif; margin: 2rem; background-color: #f7f9fc; color: #1f2937; }" +
                "h1, h2, h3 { color: #0f172a; }" +
                "header { border-bottom: 2px solid #e2e8f0; margin-bottom: 2rem; padding-bottom: 1rem; }" +
                ".meta { margin: 0.2rem 0; color: #475569; }" +
                "section { margin-bottom: 2rem; }" +
                ".chart img { max-width: 100%; height: auto; border: 1px solid #cbd5f5; background: white; padding: 1rem; }" +
                "table { width: 100%; border-collapse: collapse; margin-top: 1rem; background: white; }" +
                "th, td { border: 1px solid #e2e8f0; padding: 0.5rem 0.7rem; text-align: left; }" +
                "th { background-color: #f1f5f9; font-weight: 600; }" +
                "tr:nth-child(even) { background-color: #f8fafc; }" +
                ".sheet { background: #fff; padding: 1.5rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.1); margin-bottom: 1.5rem; }" +
                ".missing ul { list-style: disc; padding-left: 1.5rem; }" +
                ".empty { color: #64748b; font-style: italic; }";
    }
}
