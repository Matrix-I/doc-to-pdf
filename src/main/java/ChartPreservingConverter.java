import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

public class ChartPreservingConverter {

  public static void convertToPdf(String inputPath, String outputPath) {
    try {
      Document pdfDoc = new Document(PageSize.A4);
      PdfWriter.getInstance(pdfDoc, new FileOutputStream(outputPath));
      pdfDoc.open();

      FileInputStream fis = new FileInputStream(inputPath);
      XWPFDocument document = new XWPFDocument(fis);

      for (XWPFChart chart : document.getCharts()) {}

      // Process all body elements (paragraphs, tables, charts)
      for (IBodyElement element : document.getBodyElements()) {
        if (element instanceof XWPFParagraph) {
          processParagraph((XWPFParagraph) element, pdfDoc);
        } else if (element instanceof XWPFTable) {
          processTable((XWPFTable) element, pdfDoc);
        }
      }

      // Handle charts separately
      processCharts(document, pdfDoc);

      pdfDoc.close();
      document.close();
      fis.close();

      System.out.println("Conversion completed!");

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void processParagraph(XWPFParagraph paragraph, Document pdfDoc)
      throws DocumentException, IOException {
    String text = paragraph.getText();
    if (text != null && !text.trim().isEmpty()) {
      pdfDoc.add(new Paragraph(text));
    }

    // Handle images in paragraph
    for (XWPFRun run : paragraph.getRuns()) {
      List<XWPFPicture> pictures = run.getEmbeddedPictures();
      for (XWPFPicture picture : pictures) {
        addImageToPdf(picture.getPictureData().getData(), pdfDoc);
      }
    }

    //    // Handle charts in the document
    //    XWPFDocument document = paragraph.getDocument();
    //    if (document != null) {
    //      for (XWPFChart chart : document.getCharts()) {
    //        // Check if the chart is associated with this paragraph
    //        byte[] chartImageData = convertChartToImage(chart);
    //        addImageToPdf(chartImageData, pdfDoc);
    //      }
    //    }
  }

  private static boolean isChartInParagraph(XWPFChart chart, XWPFParagraph paragraph) {
    // Check if the chart is referenced in the paragraph's runs
    for (XWPFRun run : paragraph.getRuns()) {
      if (run.getCTR().getDrawingList() != null) {
        for (CTDrawing drawing : run.getCTR().getDrawingList()) {
          for (CTInline inline : drawing.getInlineList()) {
            if (inline.getGraphic() != null && inline.getGraphic().getGraphicData() != null) {
              String uri = inline.getGraphic().getGraphicData().getUri();
              if ("http://schemas.openxmlformats.org/drawingml/2006/chart".equals(uri)) {
                String chartRelId =
                    inline
                        .getGraphic()
                        .getGraphicData()
                        .getDomNode()
                        .getAttributes()
                        .getNamedItem("r:id")
                        .getNodeValue();
                if (chartRelId != null) {
                  return true;
                }
              }
            }
          }
        }
      }
    }
    return false;
  }

  private static byte[] convertChartToImage(XWPFChart chart) throws IOException {
    // Placeholder for chart-to-image conversion
    // Use Apache POI's chart data to render as image (e.g., with Java 2D or Apache Batik)
    BufferedImage chartImage = new BufferedImage(600, 400, BufferedImage.TYPE_INT_ARGB);
    Graphics2D g2d = chartImage.createGraphics();
    // Render chart using chart.getChartSeries() or other chart data
    g2d.dispose();
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    ImageIO.write(chartImage, "PNG", baos);
    return baos.toByteArray();
  }

  private static byte[] convertChartToImage(CTChart ctChart) throws IOException {
    // This is a placeholder for chart-to-image conversion
    // Actual implementation depends on your chart rendering approach
    // You might need a library like Apache Batik or a Java graphics library to render the chart
    // For example, you could use Java 2D to draw the chart and convert to a byte array

    //    // Sample pseudocode (replace with actual rendering logic):
    BufferedImage chartImage = new BufferedImage(600, 400, BufferedImage.TYPE_INT_ARGB);
    Graphics2D g2d = chartImage.createGraphics();
    // Render chart using ctChart data (e.g., using Apache Batik or custom rendering)
    g2d.dispose();
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    ImageIO.write(chartImage, "PNG", baos);
    return baos.toByteArray();
  }

  private static void processTable(XWPFTable table, Document pdfDoc) throws DocumentException {
    PdfPTable pdfTable = new PdfPTable(table.getRow(0).getTableCells().size());

    for (XWPFTableRow row : table.getRows()) {
      for (XWPFTableCell cell : row.getTableCells()) {
        pdfTable.addCell(cell.getText());
      }
    }
    pdfDoc.add(pdfTable);
  }

  private static void processCharts(XWPFDocument document, Document pdfDoc) {
    try {
      // Note: Chart extraction is complex and may require additional libraries
      // This is a placeholder for chart processing
      pdfDoc.add(
          new Paragraph("\n[Charts from original document - complex conversion required]\n"));
    } catch (Exception e) {
      System.out.println("Chart processing failed: " + e.getMessage());
    }
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
}
