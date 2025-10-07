package com.example.reporting.email;

import com.example.reporting.config.AppConfig;
import com.example.reporting.model.ReportSummary;

import java.io.IOException;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.Properties;

import jakarta.activation.DataHandler;
import jakarta.activation.FileDataSource;
import jakarta.mail.Authenticator;
import jakarta.mail.Message;
import jakarta.mail.MessagingException;
import jakarta.mail.PasswordAuthentication;
import jakarta.mail.Session;
import jakarta.mail.Transport;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class EmailService {
    private static final Logger LOGGER = LoggerFactory.getLogger(EmailService.class);
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("MMMM dd, yyyy", Locale.US);

    private final AppConfig config;

    public EmailService(AppConfig config) {
        this.config = config;
    }

    public void sendReportEmail(ReportSummary summary, Path htmlReport, Path pdfReport) {
        if (!config.isEmailEnabled()) {
            LOGGER.info("Email sending is disabled. Skipping notification.");
            return;
        }

        List<String> recipients = config.getEmailRecipients();
        if (recipients.isEmpty()) {
            LOGGER.warn("Email is enabled but no recipients are configured.");
            return;
        }

        try {
            Session session = createSession();
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(config.getEmailFrom()));
            for (String recipient : recipients) {
                message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
            }
            message.setSubject(config.getEmailSubject());

            MimeBodyPart summaryPart = new MimeBodyPart();
            summaryPart.setContent(buildEmailBody(summary), "text/html; charset=UTF-8");

            MimeMultipart multipart = new MimeMultipart();
            multipart.addBodyPart(summaryPart);

            if (htmlReport != null) {
                multipart.addBodyPart(createAttachment(htmlReport));
            }
            if (pdfReport != null) {
                multipart.addBodyPart(createAttachment(pdfReport));
            }

            message.setContent(multipart);
            Transport.send(message);
            LOGGER.info("Report email sent to {}", String.join(", ", recipients));
        } catch (MessagingException | IOException e) {
            LOGGER.error("Failed to send report email", e);
        }
    }

    private Session createSession() {
        Properties properties = new Properties();
        properties.put("mail.smtp.host", config.getSmtpHost());
        properties.put("mail.smtp.port", String.valueOf(config.getSmtpPort()));
        properties.put("mail.smtp.starttls.enable", "true");

        String username = config.getSmtpUsername();
        String password = config.getSmtpPassword();
        boolean useAuth = username != null && !username.isBlank();
        if (useAuth) {
            properties.put("mail.smtp.auth", "true");
            return Session.getInstance(properties, new Authenticator() {
                @Override
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(username, password);
                }
            });
        }
        return Session.getInstance(properties);
    }

    private MimeBodyPart createAttachment(Path file) throws MessagingException, IOException {
        MimeBodyPart attachmentPart = new MimeBodyPart();
        FileDataSource dataSource = new FileDataSource(file.toFile());
        attachmentPart.setDataHandler(new DataHandler(dataSource));
        attachmentPart.setFileName(file.getFileName().toString());
        attachmentPart.setDisposition(MimeBodyPart.ATTACHMENT);
        return attachmentPart;
    }

    private String buildEmailBody(ReportSummary summary) {
        StringBuilder builder = new StringBuilder();
        builder.append("<p>Hello,</p>")
                .append("<p>The weekly SD statistics report has been generated.</p>")
                .append(String.format("<p><strong>Report Date:</strong> %s</p>", DATE_FORMATTER.format(summary.getReportDate())))
                .append(String.format("<p><strong>Total Rows Processed:</strong> %,d</p>", summary.getTotalRowCount()));

        if (!summary.getMissingSheets().isEmpty()) {
            builder.append("<p><strong>Missing Worksheets:</strong></p><ul>");
            summary.getMissingSheets().forEach(sheet ->
                    builder.append(String.format("<li>%s</li>", sheet)));
            builder.append("</ul>");
        }

        builder.append("<p>Please find the full HTML and PDF versions attached.</p>")
                .append("<p>Regards,<br/>SD Reporting Automation</p>");
        return builder.toString();
    }
}
