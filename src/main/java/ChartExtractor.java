import java.io.*;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

public class ChartExtractor {

  public static void main(String[] args) {
    extractChartsFromDocx("file-sample.docx");
  }

  public static void extractChartsFromDocx(String inputPath) {
    try {
      FileInputStream fis = new FileInputStream(inputPath);
      XWPFDocument document = new XWPFDocument(fis);

      System.out.println("=== Extracting Charts ===");

      // Method 1: Check paragraphs for drawings
      extractChartsFromParagraphs(document);

      // Method 2: Check document relations for chart parts
      extractChartsFromRelations(document);

      // Method 3: Check headers and footers
      extractChartsFromHeadersFooters(document);

      document.close();
      fis.close();

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void extractChartsFromParagraphs(XWPFDocument document) {
    System.out.println("Checking paragraphs for charts...");

    for (XWPFParagraph paragraph : document.getParagraphs()) {
      for (XWPFRun run : paragraph.getRuns()) {
        // Check for drawings in runs
        CTDrawing[] drawings = run.getCTR().getDrawingArray();
        if (drawings != null && drawings.length > 0) {
          System.out.println("Found " + drawings.length + " drawing(s) in paragraph");

          for (CTDrawing drawing : drawings) {
            processDrawing(drawing, document);
          }
        }
      }
    }
  }

  private static void extractChartsFromRelations(XWPFDocument document) {
    System.out.println("Checking document relations for charts...");

    try {
      PackageRelationshipCollection chartRels =
          document
              .getPackagePart()
              .getRelationshipsByType(
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");

      System.out.println("Found " + chartRels.size() + " chart relationship(s)");

      for (PackageRelationship rel : chartRels) {
        PackagePart chartPart = document.getPackagePart().getRelatedPart(rel);
        System.out.println("Chart part: " + chartPart.getPartName());

        // You can process the chart XML here
        try (InputStream chartStream = chartPart.getInputStream()) {
          // Read chart data
          byte[] chartData = chartStream.readAllBytes();
          System.out.println("Chart data size: " + chartData.length + " bytes");

          // Save chart XML for analysis
          saveChartXml(chartData, rel.getId());
        }
      }

    } catch (Exception e) {
      System.out.println("Error extracting charts from relations: " + e.getMessage());
    }
  }

  private static void extractChartsFromHeadersFooters(XWPFDocument document) {
    System.out.println("Checking headers and footers for charts...");

    // Check headers
    for (XWPFHeader header : document.getHeaderList()) {
      for (XWPFParagraph paragraph : header.getParagraphs()) {
        for (XWPFRun run : paragraph.getRuns()) {
          CTDrawing[] drawings = run.getCTR().getDrawingArray();
          if (drawings != null && drawings.length > 0) {
            System.out.println("Found drawing(s) in header");
          }
        }
      }
    }

    // Check footers
    for (XWPFFooter footer : document.getFooterList()) {
      for (XWPFParagraph paragraph : footer.getParagraphs()) {
        for (XWPFRun run : paragraph.getRuns()) {
          CTDrawing[] drawings = run.getCTR().getDrawingArray();
          if (drawings != null && drawings.length > 0) {
            System.out.println("Found drawing(s) in footer");
          }
        }
      }
    }
  }

  private static void processDrawing(CTDrawing drawing, XWPFDocument document) {
    System.out.println("Processing drawing...");
    // This is where you would extract chart information
    // The drawing object contains references to charts and other graphics
  }

  private static void saveChartXml(byte[] chartData, String chartId) {
    try {
      FileOutputStream fos = new FileOutputStream("chart_" + chartId + ".xml");
      fos.write(chartData);
      fos.close();
      System.out.println("Saved chart XML: chart_" + chartId + ".xml");
    } catch (Exception e) {
      System.out.println("Error saving chart XML: " + e.getMessage());
    }
  }
}
