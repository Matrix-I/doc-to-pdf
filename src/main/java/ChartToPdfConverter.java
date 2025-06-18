import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;

public class ChartToPdfConverter {
  public static void main(String[] args) {
    String inputPath = "file-sample.docx"; // Path to input DOCX file with chart
    String outputPath = "output111.pdf"; // Path to output PDF file

    try {
      convertChartToPdf(inputPath, outputPath);
      System.out.println("Chart conversion completed successfully!");
    } catch (Exception e) {
      System.err.println("Error during conversion: " + e.getMessage());
    }
  }

  public static void convertChartToPdf(String docxPath, String pdfPath) throws Exception {
    // Load DOCX file
    // Create PDF document
    try (FileInputStream fis = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
        PDDocument pdfDocument = new PDDocument()) {
      PDPage page = new PDPage();
      pdfDocument.addPage(page);
      PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);

      contentStream.setFont(PDType1Font.HELVETICA, 12);
      float yOffset = 700; // Starting vertical position
      float leading = 15; // Line spacing
      float margin = 50;

      // Find and process charts
      List<String> chartInfo = new ArrayList<>();
      for (XWPFParagraph paragraph : document.getParagraphs()) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        cursor.selectPath(
            "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' "
                + "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' "
                + "declare namespace c='http://schemas.openxmlformats.org/drawingml/2006/chart' "
                + ".//c:chart");

        while (cursor.hasNextSelection()) {
          cursor.toNextSelection();
          XmlObject xmlObject = cursor.getObject();
          String chartXml = xmlObject.xmlText();

          // Parse chart XML (simplified)
          if (xmlObject
              instanceof
              org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTChartImpl
              ctChart) {
            // Get chart title
            String title = "No Title";
            if (ctChart.getTitle() != null) {
              CTTitle ctTitle = ctChart.getTitle();
              if (ctTitle.getTx() != null && ctTitle.getTx().getRich() != null) {
                title = ctTitle.getTx().getRich().toString();
              }
            }
            chartInfo.add("Chart Title: " + title);

            // Get chart type (e.g., bar, line, pie)
            String chartType =
                ctChart.getPlotArea().xmlText().contains("barChart")
                    ? "Bar Chart"
                    : ctChart.getPlotArea().xmlText().contains("lineChart")
                        ? "Line Chart"
                        : ctChart.getPlotArea().xmlText().contains("pieChart")
                            ? "Pie Chart"
                            : "Unknown";
            chartInfo.add("Chart Type: " + chartType);

            // Get series data (simplified)
            if (ctChart.getPlotArea().getBarChartList() != null) {
              for (var barChart : ctChart.getPlotArea().getBarChartList()) {
                for (CTBarSer ser : barChart.getSerList()) {
                  String seriesName =
                      ser.getTx() != null && ser.getTx().getStrRef() != null
                          ? ser.getTx().getStrRef().getF()
                          : "Unnamed Series";
                  chartInfo.add("Series: " + seriesName);
                }
              }
            }
          }
        }
        cursor.dispose();
      }

      // Write chart info to PDF
      contentStream.beginText();
      contentStream.newLineAtOffset(margin, yOffset);
      contentStream.showText("Chart Information (Placeholder for /word/charts/chart1.xml)");
      contentStream.endText();
      yOffset -= leading;

      for (String info : chartInfo) {
        contentStream.beginText();
        contentStream.newLineAtOffset(margin, yOffset);
        contentStream.showText(info);
        contentStream.endText();
        yOffset -= leading;

        // Check for new page
        if (yOffset < 50) {
          contentStream.close();
          page = new PDPage();
          pdfDocument.addPage(page);
          contentStream = new PDPageContentStream(pdfDocument, page);
          contentStream.setFont(PDType1Font.HELVETICA, 12);
          yOffset = 700;
        }
      }

      contentStream.close();
      pdfDocument.save(pdfPath);
    }
  }
}
