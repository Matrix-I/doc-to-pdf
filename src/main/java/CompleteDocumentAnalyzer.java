import java.io.*;
import java.util.List;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xwpf.usermodel.*;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

public class CompleteDocumentAnalyzer {

  public static void analyzeDocument(String inputPath) {
    try {
      FileInputStream fis = new FileInputStream(inputPath);
      XWPFDocument document = new XWPFDocument(fis);

      System.out.println("=== Complete Document Analysis ===");

      // Analyze all package parts
      analyzePackageParts(document);

      // Analyze all relationships
      analyzeRelationships(document);

      // Check for embedded objects
      checkEmbeddedObjects(document);

      document.close();
      fis.close();

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void analyzePackageParts(XWPFDocument document) {
    System.out.println("\n--- Package Parts ---");

    try {
      OPCPackage pkg = document.getPackage();

      for (PackagePart part : pkg.getParts()) {
        String partName = part.getPartName().toString();
        String contentType = part.getContentType();

        System.out.println("Part: " + partName + " (Type: " + contentType + ")");

        // Check for chart parts
        if (contentType.contains("chart")) {
          System.out.println("  --> CHART FOUND!");
          extractChartData(part);
        }

        // Check for drawing parts
        if (contentType.contains("drawing")) {
          System.out.println("  --> DRAWING FOUND!");
        }

        // Check for embedded Excel files (charts often have embedded data)
        if (contentType.contains("excel") || contentType.contains("spreadsheet")) {
          System.out.println("  --> EXCEL DATA FOUND!");
        }
      }

    } catch (Exception e) {
      System.out.println("Error analyzing package parts: " + e.getMessage());
    }
  }

  private static void analyzeRelationships(XWPFDocument document) {
    System.out.println("\n--- Relationships ---");

    try {
      PackagePart mainPart = document.getPackagePart();

      for (PackageRelationship rel : mainPart.getRelationships()) {
        System.out.println("Relationship: " + rel.getRelationshipType());
        System.out.println("  Target: " + rel.getTargetURI());
        System.out.println("  ID: " + rel.getId());

        if (rel.getRelationshipType().contains("chart")) {
          System.out.println("  --> CHART RELATIONSHIP!");
        }
      }

    } catch (Exception e) {
      System.out.println("Error analyzing relationships: " + e.getMessage());
    }
  }

  private static void checkEmbeddedObjects(XWPFDocument document) {
    System.out.println("\n--- Embedded Objects ---");

    // Check all paragraphs for embedded objects
    for (XWPFParagraph paragraph : document.getParagraphs()) {
      for (XWPFRun run : paragraph.getRuns()) {

        // Check for embedded objects
        if (run.getCTR().getObjectArray() != null && run.getCTR().getObjectArray().length > 0) {
          System.out.println("Found embedded object in paragraph");
        }

        // Check for pictures (charts might be saved as images)
        List<XWPFPicture> pictures = run.getEmbeddedPictures();
        if (!pictures.isEmpty()) {
          System.out.println(
              "Found " + pictures.size() + " picture(s) - might include chart images");

          for (int i = 0; i < pictures.size(); i++) {
            XWPFPicture picture = pictures.get(i);
            String fileName = picture.getPictureData().getFileName();
            System.out.println("  Picture " + i + ": " + fileName);

            // Save picture for analysis
            savePicture(picture, i);
          }
        }
      }
    }
  }

  private static void extractChartData(PackagePart chartPart) {
    try {
      InputStream chartStream = chartPart.getInputStream();

      // Parse the chart XML
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      DocumentBuilder builder = factory.newDocumentBuilder();
      Document chartDoc = builder.parse(new InputSource(chartStream));

      System.out.println("  Chart XML parsed successfully");

      // You can now traverse the chart XML to extract data
      // This would involve parsing the chart structure, series data, etc.

      chartStream.close();

    } catch (Exception e) {
      System.out.println("  Error extracting chart data: " + e.getMessage());
    }
  }

  private static void savePicture(XWPFPicture picture, int index) {
    try {
      byte[] imageData = picture.getPictureData().getData();
      //      String extension = getImageExtension(picture.getPictureData().getPictureType());
      String extension = ".png";

      FileOutputStream fos = new FileOutputStream("extracted_image_" + index + extension);
      fos.write(imageData);
      fos.close();

      System.out.println("    Saved as: extracted_image_" + index + extension);

    } catch (Exception e) {
      System.out.println("    Error saving picture: " + e.getMessage());
    }
  }

  //  private static String getImageExtension(int pictureType) {
  //    return switch (pictureType) {
  //      case XWPFDocument.PICTURE_TYPE_JPEG -> ".jpg";
  //      case XWPFDocument.PICTURE_TYPE_PNG -> ".png";
  //      case XWPFDocument.PICTURE_TYPE_GIF -> ".gif";
  //      case XWPFDocument.PICTURE_TYPE_BMP -> ".bmp";
  //      case XWPFDocument.PICTURE_TYPE_WMF -> ".wmf";
  //      case XWPFDocument.PICTURE_TYPE_EMF -> ".emf";
  //      default -> ".unknown";
  //    };
  //  }

  public static void main(String[] args) {
    String inputFile = "file-sample.docx";
    analyzeDocument(inputFile);
  }
}
