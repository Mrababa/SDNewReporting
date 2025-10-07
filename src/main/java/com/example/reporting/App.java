package com.example.reporting;

import com.example.reporting.config.AppConfig;
import com.example.reporting.data.ExcelReportReader;
import com.example.reporting.email.EmailService;
import com.example.reporting.file.ReportFileLocator;
import com.example.reporting.model.ReportFile;
import com.example.reporting.model.ReportSummary;
import com.example.reporting.report.ReportGenerator;
import com.example.reporting.util.ChartRenderer;

import java.nio.file.Path;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class App {
    private static final Logger LOGGER = LoggerFactory.getLogger(App.class);
    private static final List<String> EXPECTED_SHEETS = List.of(
            "VW_Abnormal_IDs",
            "VW_ICP_ApiSe_Stats",
            "VW_ICPSeErrorsDetails",
            "VW_MemUploadTCount",
            "VW_SD_SeErrorDetails",
            "VW_SD_SeErrorDetailsIC",
            "VW_SD_SeHitCount"
    );

    public static void main(String[] args) {
        App app = new App();
        try {
            app.run();
        } catch (Exception e) {
            LOGGER.error("Failed to generate SD report", e);
            System.exit(1);
        }
    }

    private void run() {
        AppConfig config = new AppConfig();
        ReportFileLocator locator = new ReportFileLocator(config.getReportDirectory(), config.getReportFilePattern());
        ReportFile reportFile = locator.findLatestReportFile()
                .orElseThrow(() -> new IllegalStateException("No report files were found in " + config.getReportDirectory()));

        LOGGER.info("Processing report file: {}", reportFile.path());

        ExcelReportReader reader = new ExcelReportReader(EXPECTED_SHEETS);
        ReportSummary summary = reader.readReport(reportFile);

        ReportGenerator generator = new ReportGenerator(new ChartRenderer());
        String htmlContent = generator.buildHtmlReport(summary);

        String baseFileName = String.format("SD_Weekly_Summary_%s", summary.getReportDate());
        Path outputDirectory = config.getOutputDirectory();
        Path htmlReport = generator.writeHtmlReport(outputDirectory, baseFileName + ".html", htmlContent);
        Path pdfReport = generator.writePdfReport(outputDirectory, baseFileName + ".pdf", htmlContent);

        EmailService emailService = new EmailService(config);
        emailService.sendReportEmail(summary, htmlReport, pdfReport);
    }
}
