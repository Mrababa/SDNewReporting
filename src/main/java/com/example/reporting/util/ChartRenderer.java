package com.example.reporting.util;

import com.example.reporting.model.SheetSummary;

import java.awt.Color;
import java.awt.BasicStroke;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.Map;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ChartRenderer {
    private static final Logger LOGGER = LoggerFactory.getLogger(ChartRenderer.class);

    public String renderRowCountChart(Map<String, SheetSummary> sheetSummaries) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        sheetSummaries.values().forEach(summary ->
                dataset.addValue(summary.getRowCount(), "Rows", summary.getSheetName()));

        JFreeChart chart = ChartFactory.createBarChart(
                "Rows per Sheet",
                "Sheet",
                "Row Count",
                dataset,
                PlotOrientation.VERTICAL,
                false,
                true,
                false);

        styleChart(chart);

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            ChartUtils.writeChartAsPNG(outputStream, chart, 900, 500);
            return Base64.getEncoder().encodeToString(outputStream.toByteArray());
        } catch (IOException e) {
            LOGGER.error("Failed to render chart", e);
            return "";
        }
    }

    private void styleChart(JFreeChart chart) {
        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setDomainGridlinePaint(Color.LIGHT_GRAY);
        plot.setRangeGridlinePaint(Color.LIGHT_GRAY);
        plot.getRenderer().setSeriesPaint(0, new Color(79, 129, 189));
        plot.getRenderer().setDefaultStroke(new BasicStroke(2.0f));
    }
}
