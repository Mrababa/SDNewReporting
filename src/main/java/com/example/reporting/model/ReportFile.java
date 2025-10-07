package com.example.reporting.model;

import java.nio.file.Path;
import java.time.LocalDate;

public record ReportFile(Path path, LocalDate reportDate) {
}
