import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

public class XWPFChartToImageConverter {

  /** Convert XWPFChart to BufferedImage using JFreeChart */
  public static BufferedImage convertChartToImage(XWPFChart chart, int width, int height) {
    try {
      CTChart ctChart = chart.getCTChart();

      // Determine chart type and create appropriate JFreeChart
      JFreeChart jFreeChart = createJFreeChartFromCTChart(ctChart);

      if (jFreeChart != null) {
        return jFreeChart.createBufferedImage(width, height);
      } else {
        return createFallbackChartImage(chart, width, height);
      }

    } catch (Exception e) {
      System.err.println("Error converting chart: " + e.getMessage());
      return createFallbackChartImage(chart, width, height);
    }
  }

  /** Create JFreeChart from CTChart based on chart type */
  private static JFreeChart createJFreeChartFromCTChart(CTChart ctChart) {
    try {
      // Get chart title
      String title = "";
      if (ctChart.getTitle() != null && ctChart.getTitle().getTx() != null) {
        if (ctChart.getTitle().getTx().getRich() != null) {
          title = extractTextFromRich(ctChart.getTitle().getTx().getRich());
        }
      }

      // Check for different chart types
      CTPlotArea plotArea = ctChart.getPlotArea();

      // Bar Chart
      if (plotArea.getBarChartArray().length > 0) {
        return createBarChart(plotArea.getBarChartArray(0), title);
      }

      // Line Chart
      if (plotArea.getLineChartArray().length > 0) {
        return createLineChart(plotArea.getLineChartArray(0), title);
      }

      // Pie Chart
      if (plotArea.getPieChartArray().length > 0) {
        return createPieChart(plotArea.getPieChartArray(0), title);
      }

      // Area Chart
      if (plotArea.getAreaChartArray().length > 0) {
        return createAreaChart(plotArea.getAreaChartArray(0), title);
      }

    } catch (Exception e) {
      System.err.println("Error creating JFreeChart: " + e.getMessage());
    }

    return null;
  }

  /** Create Bar Chart from CTBarChart */
  private static JFreeChart createBarChart(CTBarChart barChart, String title) {
    DefaultCategoryDataset dataset = new DefaultCategoryDataset();

    try {
      CTBarSer[] series = barChart.getSerArray();

      for (CTBarSer ser : series) {
        String seriesName = getSeriesName(ser.getTx());

        // Get categories (X-axis labels)
        String[] categories = getCategories(ser.getCat());

        // Get values (Y-axis values)
        double[] values = getValues(ser.getVal());

        // Add data to dataset
        for (int i = 0; i < Math.min(categories.length, values.length); i++) {
          dataset.addValue(values[i], seriesName, categories[i]);
        }
      }

      return ChartFactory.createBarChart(
          title, "Category", "Value", dataset, PlotOrientation.VERTICAL, true, true, false);

    } catch (Exception e) {
      System.err.println("Error creating bar chart: " + e.getMessage());
    }

    return null;
  }

  /** Create Line Chart from CTLineChart */
  private static JFreeChart createLineChart(CTLineChart lineChart, String title) {
    DefaultCategoryDataset dataset = new DefaultCategoryDataset();

    try {
      CTLineSer[] series = lineChart.getSerArray();

      for (CTLineSer ser : series) {
        String seriesName = getSeriesName(ser.getTx());

        // Get categories and values
        String[] categories = getCategories(ser.getCat());
        double[] values = getValues(ser.getVal());

        // Add data to dataset
        for (int i = 0; i < Math.min(categories.length, values.length); i++) {
          dataset.addValue(values[i], seriesName, categories[i]);
        }
      }

      return ChartFactory.createLineChart(
          title, "Category", "Value", dataset, PlotOrientation.VERTICAL, true, true, false);

    } catch (Exception e) {
      System.err.println("Error creating line chart: " + e.getMessage());
    }

    return null;
  }

  /** Create Pie Chart from CTPieChart */
  private static JFreeChart createPieChart(CTPieChart pieChart, String title) {
    DefaultPieDataset dataset = new DefaultPieDataset();

    try {
      CTPieSer[] series = pieChart.getSerArray();

      if (series.length > 0) {
        CTPieSer ser = series[0]; // Usually only one series in pie chart

        // Get categories and values
        String[] categories = getCategories(ser.getCat());
        double[] values = getValues(ser.getVal());

        // Add data to dataset
        for (int i = 0; i < Math.min(categories.length, values.length); i++) {
          dataset.setValue(categories[i], values[i]);
        }
      }

      return ChartFactory.createPieChart(title, dataset, true, true, false);

    } catch (Exception e) {
      System.err.println("Error creating pie chart: " + e.getMessage());
    }

    return null;
  }

  /** Create Area Chart from CTAreaChart */
  private static JFreeChart createAreaChart(CTAreaChart areaChart, String title) {
    DefaultCategoryDataset dataset = new DefaultCategoryDataset();

    try {
      CTAreaSer[] series = areaChart.getSerArray();

      for (CTAreaSer ser : series) {
        String seriesName = getSeriesName(ser.getTx());

        // Get categories and values
        String[] categories = getCategories(ser.getCat());
        double[] values = getValues(ser.getVal());

        // Add data to dataset
        for (int i = 0; i < Math.min(categories.length, values.length); i++) {
          dataset.addValue(values[i], seriesName, categories[i]);
        }
      }

      return ChartFactory.createAreaChart(
          title, "Category", "Value", dataset, PlotOrientation.VERTICAL, true, true, false);

    } catch (Exception e) {
      System.err.println("Error creating area chart: " + e.getMessage());
    }

    return null;
  }

  /** Extract series name from CTSerTx */
  private static String getSeriesName(CTSerTx serTx) {
    if (serTx != null) {
      if (serTx.getStrRef() != null && serTx.getStrRef().getStrCache() != null) {
        CTStrData strData = serTx.getStrRef().getStrCache();
        if (strData.getPtArray().length > 0) {
          return strData.getPtArray(0).getV();
        }
      } else if (serTx.getV() != null) {
        return serTx.getV();
      }
    }
    return "Series";
  }

  /** Extract categories from CTAxDataSource */
  private static String[] getCategories(CTAxDataSource cat) {
    try {
      if (cat != null && cat.getStrRef() != null && cat.getStrRef().getStrCache() != null) {
        CTStrData strData = cat.getStrRef().getStrCache();
        CTStrVal[] points = strData.getPtArray();

        String[] categories = new String[points.length];
        for (int i = 0; i < points.length; i++) {
          categories[i] = points[i].getV();
        }
        return categories;
      }
    } catch (Exception e) {
      System.err.println("Error extracting categories: " + e.getMessage());
    }

    return new String[] {"Category 1", "Category 2", "Category 3"};
  }

  /** Extract values from CTNumDataSource */
  private static double[] getValues(CTNumDataSource val) {
    try {
      if (val != null && val.getNumRef() != null && val.getNumRef().getNumCache() != null) {
        CTNumData numData = val.getNumRef().getNumCache();
        CTNumVal[] points = numData.getPtArray();

        double[] values = new double[points.length];
        for (int i = 0; i < points.length; i++) {
          try {
            values[i] = Double.parseDouble(points[i].getV());
          } catch (NumberFormatException e) {
            values[i] = 0.0;
          }
        }
        return values;
      }
    } catch (Exception e) {
      System.err.println("Error extracting values: " + e.getMessage());
    }

    return new double[] {10.0, 20.0, 15.0};
  }

  /** Extract text from rich text format */
  private static String extractTextFromRich(
      org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody rich) {
    try {
      if (rich.getPArray().length > 0) {
        org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph p = rich.getPArray(0);
        if (p.getRArray().length > 0) {
          return p.getRArray(0).getT();
        }
      }
    } catch (Exception e) {
      System.err.println("Error extracting rich text: " + e.getMessage());
    }
    return "";
  }

  /** Create fallback image when chart conversion fails */
  private static BufferedImage createFallbackChartImage(XWPFChart chart, int width, int height) {
    BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
    Graphics2D g2d = image.createGraphics();

    // Set rendering hints
    g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
    g2d.setRenderingHint(
        RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

    // Fill background
    g2d.setColor(Color.WHITE);
    g2d.fillRect(0, 0, width, height);

    // Draw border
    g2d.setColor(Color.GRAY);
    g2d.drawRect(10, 10, width - 20, height - 20);

    // Draw title
    g2d.setColor(Color.BLACK);
    g2d.setFont(new Font("Arial", Font.BOLD, 18));
    String title = "Chart from DOCX";

    try {
      // Try to get actual chart title
      CTChart ctChart = chart.getCTChart();
      if (ctChart.getTitle() != null && ctChart.getTitle().getTx() != null) {
        if (ctChart.getTitle().getTx().getRich() != null) {
          String extractedTitle = extractTextFromRich(ctChart.getTitle().getTx().getRich());
          if (!extractedTitle.isEmpty()) {
            title = extractedTitle;
          }
        }
      }
    } catch (Exception e) {
      // Use default title
    }

    int titleWidth = g2d.getFontMetrics().stringWidth(title);
    g2d.drawString(title, (width - titleWidth) / 2, 50);

    // Draw chart info
    g2d.setFont(new Font("Arial", Font.PLAIN, 14));
    g2d.drawString("Chart extracted from DOCX document", 30, height - 60);
    g2d.drawString("Chart rendering requires data extraction", 30, height - 40);
    g2d.drawString("Consider using JFreeChart for full rendering", 30, height - 20);

    g2d.dispose();
    return image;
  }

  /** Main method for testing */
  public static void main(String[] args) {
    try {
      String docxPath = "file-sample.docx";
      String outputDir = "";

      // Create output directory
      new File(outputDir).mkdirs();

      try (FileInputStream fis = new FileInputStream(docxPath);
          XWPFDocument document = new XWPFDocument(fis)) {

        List<XWPFChart> charts = document.getCharts();
        System.out.println("Found " + charts.size() + " charts");

        for (int i = 0; i < charts.size(); i++) {
          XWPFChart chart = charts.get(i);
          BufferedImage chartImage = convertChartToImage(chart, 800, 600);

          String fileName = outputDir + "/chart_" + (i + 1) + ".png";
          ImageIO.write(chartImage, "PNG", new File(fileName));
          System.out.println("Saved chart: " + fileName);
        }
      }

    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
