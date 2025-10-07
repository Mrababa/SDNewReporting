# SDNewReporting

This project automates the collection of weekly SD statistics contained in Excel workbooks and produces human-friendly summary reports.

## Features

* Discovers weekly Excel workbooks that follow the `StatsReports_YYYYMMDD.xlsx` naming convention.
* Validates the presence of all required worksheets and highlights any missing tabs.
* Extracts per-sheet statistics such as row counts, column headers, and representative sample rows.
* Builds a styled HTML summary that includes tabular details and an automatically generated bar chart of worksheet volumes.
* Produces a PDF version of the same report using the HTML rendering pipeline.
* Sends configurable email notifications with the HTML and PDF reports attached.

## Configuration

Update `src/main/resources/config.properties` to match your environment:

```
report.directory=reports                 # Folder that stores incoming Excel workbooks
output.directory=generated-reports       # Folder where HTML/PDF reports are written
report.file.pattern=StatsReports_yyyyMMdd.xlsx
email.enabled=false                      # Flip to true to enable email notifications
email.recipients=team@example.com        # Comma-separated list of recipients
email.subject=Weekly SD Stats Summary
email.from=reports@example.com
smtp.host=smtp.example.com
smtp.port=587
smtp.username=user@example.com
smtp.password=changeme
```

When email is enabled the application will use the SMTP settings above to send the summary message. Leave `email.enabled=false` for local testing.

## Running the Application

Build and execute the program with Maven:

```
mvn clean package
java -cp target/hello-world-1.0-SNAPSHOT.jar com.example.reporting.App
```

Ensure the configured `report.directory` contains at least one Excel workbook that matches the naming pattern before running the application.
