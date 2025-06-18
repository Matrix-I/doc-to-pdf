import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.*;
import java.util.List;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xwpf.usermodel.*;

public class PracticalChartConverter {

  public static void convertWithChartDetection(String inputPath, String outputPath) {
    try {
      Document pdfDoc = new Document();
      PdfWriter.getInstance(pdfDoc, new FileOutputStream(outputPath));
      pdfDoc.open();

      FileInputStream fis = new FileInputStream(inputPath);
      XWPFDocument document = new XWPFDocument(fis);

      // Process regular content
      processDocumentContent(document, pdfDoc);

      // Detect and handle charts
      detectAndProcessCharts(document, pdfDoc);

      pdfDoc.close();
      document.close();
      fis.close();

      System.out.println("Conversion completed with chart detection!");

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void processDocumentContent(XWPFDocument document, Document pdfDoc)
      throws DocumentException {

    for (XWPFParagraph paragraph : document.getParagraphs()) {
      String text = paragraph.getText();
      if (text != null && !text.trim().isEmpty()) {
        pdfDoc.add(new Paragraph(text));
      }

      // Handle images in paragraphs
      for (XWPFRun run : paragraph.getRuns()) {
        List<XWPFPicture> pictures = run.getEmbeddedPictures();
        for (XWPFPicture picture : pictures) {
          addImageToPdf(picture.getPictureData().getData(), pdfDoc);
        }
      }
    }

    // Process tables
    for (XWPFTable table : document.getTables()) {
      processTable(table, pdfDoc);
    }
  }

  private static void detectAndProcessCharts(XWPFDocument document, Document pdfDoc)
      throws DocumentException {

    try {
      PackageRelationshipCollection chartRels =
          document
              .getPackagePart()
              .getRelationshipsByType(
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");

      if (!chartRels.isEmpty()) {
        pdfDoc.add(new Paragraph("\n=== CHARTS DETECTED ==="));
        pdfDoc.add(new Paragraph("Found " + chartRels.size() + " chart(s) in the document."));

        for (int i = 0; i < chartRels.size(); i++) {
          pdfDoc.add(
              new Paragraph(
                  "Chart " + (i + 1) + ": [Chart content - requires specialized conversion]"));
        }

        pdfDoc.add(
            new Paragraph(
                "Note: For full chart preservation, use LibreOffice or Aspose.Words conversion."));
        pdfDoc.add(new Paragraph("=========================\n"));
      }

    } catch (Exception e) {
      System.out.println("Error detecting charts: " + e.getMessage());
    }
  }

  private static void processTable(XWPFTable table, Document pdfDoc) throws DocumentException {
    com.itextpdf.text.pdf.PdfPTable pdfTable =
        new com.itextpdf.text.pdf.PdfPTable(table.getRow(0).getTableCells().size());

    for (XWPFTableRow row : table.getRows()) {
      for (XWPFTableCell cell : row.getTableCells()) {
        pdfTable.addCell(cell.getText());
      }
    }
    pdfDoc.add(pdfTable);
  }

  private static void addImageToPdf(byte[] imageData, Document pdfDoc) {
    try {
      Image image = Image.getInstance(imageData);
      if (image.getWidth() > 500) {
        image.scaleToFit(500, 400);
      }
      pdfDoc.add(image);
    } catch (Exception e) {
      System.out.println("Could not add image: " + e.getMessage());
    }
  }

  public static void main(String[] args) {
    String inputFile = "file-sample.docx";
    String outputFile = "output_with_chart_detection.pdf";

    convertWithChartDetection(inputFile, outputFile);
  }
}
