package com.example.reporting.config;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AppConfig {
    private static final Logger LOGGER = LoggerFactory.getLogger(AppConfig.class);
    private final Properties properties = new Properties();

    public AppConfig() {
        loadProperties("config.properties");
    }

    public void loadProperties(String resourceName) {
        try (InputStream inputStream = getClass().getClassLoader().getResourceAsStream(resourceName)) {
            if (inputStream == null) {
                throw new IllegalStateException("Unable to find configuration file: " + resourceName);
            }
            properties.load(inputStream);
        } catch (IOException e) {
            throw new IllegalStateException("Failed to load configuration file: " + resourceName, e);
        }
    }

    public Path getReportDirectory() {
        String directory = properties.getProperty("report.directory", "reports");
        Path path = Paths.get(directory).toAbsolutePath().normalize();
        ensureDirectoryExists(path);
        return path;
    }

    public Path getOutputDirectory() {
        String directory = properties.getProperty("output.directory", "generated-reports");
        Path path = Paths.get(directory).toAbsolutePath().normalize();
        ensureDirectoryExists(path);
        return path;
    }

    private void ensureDirectoryExists(Path path) {
        if (!Files.exists(path)) {
            try {
                Files.createDirectories(path);
                LOGGER.info("Created directory: {}", path);
            } catch (IOException e) {
                throw new IllegalStateException("Unable to create directory: " + path, e);
            }
        }
    }

    public DateTimeFormatter getReportFileDateFormatter() {
        String pattern = properties.getProperty("report.file.pattern", "StatsReports_yyyyMMdd.xlsx");
        String datePattern = pattern.replace("StatsReports_", "").replace(".xlsx", "");
        return DateTimeFormatter.ofPattern(datePattern);
    }

    public String getReportFilePattern() {
        return properties.getProperty("report.file.pattern", "StatsReports_yyyyMMdd.xlsx");
    }

    public boolean isEmailEnabled() {
        return Boolean.parseBoolean(properties.getProperty("email.enabled", "false"));
    }

    public List<String> getEmailRecipients() {
        String recipients = properties.getProperty("email.recipients", "");
        if (recipients.isBlank()) {
            return List.of();
        }
        return Arrays.stream(recipients.split(","))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .toList();
    }

    public String getEmailSubject() {
        return properties.getProperty("email.subject", "Weekly SD Stats Summary");
    }

    public String getEmailFrom() {
        return properties.getProperty("email.from", "reports@example.com");
    }

    public String getSmtpHost() {
        return properties.getProperty("smtp.host");
    }

    public int getSmtpPort() {
        return Integer.parseInt(properties.getProperty("smtp.port", "587"));
    }

    public String getSmtpUsername() {
        return properties.getProperty("smtp.username");
    }

    public String getSmtpPassword() {
        return properties.getProperty("smtp.password");
    }
}
