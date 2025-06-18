import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

public class DocxChartToPdfConverter {

  public static void main(String[] args) {
    try {
      String docxPath = "file-sample.docx";
      String pdfPath = "output123.pdf";

      convertDocxChartsToPdf(docxPath, pdfPath);
      System.out.println("Conversion completed successfully!");

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  public static void convertDocxChartsToPdf(String docxPath, String pdfPath) throws Exception {
    // Extract charts and images from DOCX
    List<byte[]> imageData = extractImagesFromDocx(docxPath);

    // Create PDF with extracted images
    createPdfWithImages(imageData, pdfPath);
  }

  private static List<byte[]> extractImagesFromDocx(String docxPath) throws Exception {
    List<byte[]> imageDataList = new ArrayList<>();

    try (FileInputStream fis = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fis)) {

      // Method 1: Extract all pictures (including charts saved as images)
      List<XWPFPictureData> pictures = document.getAllPictures();
      for (XWPFPictureData picture : pictures) {
        byte[] data = picture.getData();
        if (data != null && data.length > 0) {
          imageDataList.add(data);
        }
      }

      // Method 2: Try to extract chart images specifically
      List<XWPFChart> charts = document.getCharts();
      for (XWPFChart chart : charts) {
        // Create a placeholder image for charts that can't be directly extracted
        BufferedImage chartImage = createChartPlaceholder(chart, imageDataList.size() + 1);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(chartImage, "PNG", baos);
        imageDataList.add(baos.toByteArray());
      }
    }

    return imageDataList;
  }

  private static BufferedImage createChartPlaceholder(XWPFChart chart, int chartNumber) {
    int width = 800;
    int height = 600;

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
    g2d.setFont(new java.awt.Font("Arial", java.awt.Font.BOLD, 24));
    String title = "Chart " + chartNumber;
    g2d.drawString(title, 50, 60);

    // Draw chart info if available
    g2d.setFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 16));
    g2d.drawString("Chart extracted from DOCX", 50, 100);

    try {
      // Try to get some basic chart information
      String chartTitle =
          chart.getCTChart().getTitle() != null
              ? chart.getCTChart().getTitle().toString()
              : "Untitled Chart";
      g2d.drawString("Title: " + chartTitle, 50, 130);
    } catch (Exception e) {
      g2d.drawString("Chart data available", 50, 130);
    }

    g2d.dispose();
    return image;
  }

  private static void createPdfWithImages(List<byte[]> imageDataList, String pdfPath)
      throws Exception {
    PdfWriter writer = new PdfWriter(pdfPath);
    PdfDocument pdfDoc = new PdfDocument(writer);
    Document document = new Document(pdfDoc);

    // Add title
    document.add(new Paragraph("Charts and Images from DOCX Document").setBold().setFontSize(16));

    // Add each image to the PDF
    for (int i = 0; i < imageDataList.size(); i++) {
      byte[] imageBytes = imageDataList.get(i);

      try {
        // Create iText Image
        ImageData imageData = ImageDataFactory.create(imageBytes);
        Image image = new Image(imageData);

        // Scale image to fit page if necessary
        float maxWidth = document.getPdfDocument().getDefaultPageSize().getWidth() - 80;
        float maxHeight = document.getPdfDocument().getDefaultPageSize().getHeight() - 150;

        if (image.getImageWidth() > maxWidth || image.getImageHeight() > maxHeight) {
          image.scaleToFit(maxWidth, maxHeight);
        }

        // Add image title
        document.add(new Paragraph("Image/Chart " + (i + 1)).setBold().setFontSize(12));

        // Add image to document
        document.add(image);

        // Add some space between images
        document.add(new Paragraph("\n"));

      } catch (Exception e) {
        // If image can't be processed, add error message
        document.add(
            new Paragraph("Could not process image " + (i + 1) + ": " + e.getMessage())
                .setItalic());
      }
    }

    if (imageDataList.isEmpty()) {
      document.add(new Paragraph("No images or charts found in the DOCX document."));
    }

    document.close();
  }

  // Enhanced method for better chart extraction using POI's chart API
  public static void extractChartsAdvanced(String docxPath, String outputDir) throws Exception {
    try (FileInputStream fis = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fis)) {

      List<XWPFChart> charts = document.getCharts();

      for (int i = 0; i < charts.size(); i++) {
        XWPFChart chart = charts.get(i);

        // Access the underlying chart data
        String chartXml = chart.getCTChart().toString();

        // You can process the chart XML to extract data and recreate the chart
        // using libraries like JFreeChart or similar

        // For demonstration, save the chart XML
        String chartFileName = outputDir + "/chart_" + i + ".xml";
        try (FileWriter writer = new FileWriter(chartFileName)) {
          writer.write(chartXml);
        }

        System.out.println("Chart XML saved: " + chartFileName);
      }
    }
  }

  // Method to create charts from extracted data using JFreeChart (optional)
  public static BufferedImage createChartFromData(String chartXml) {
    // This would require parsing the chart XML and using JFreeChart
    // to recreate the chart as an image
    // Implementation depends on your specific chart types and requirements

    // Placeholder implementation
    BufferedImage image = new BufferedImage(800, 600, BufferedImage.TYPE_INT_RGB);
    Graphics2D g2d = image.createGraphics();
    g2d.setColor(java.awt.Color.WHITE);
    g2d.fillRect(0, 0, 800, 600);
    g2d.setColor(java.awt.Color.BLACK);
    g2d.drawString("Chart placeholder", 50, 50);
    g2d.dispose();

    return image;
  }
}
