import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.*;
import java.util.List;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.xwpf.usermodel.*;

public class AdvancedDocToPdfConverter {

  public static void convertToPdf(String inputPath, String outputPath) {
    try {
      Document pdfDoc = new Document();
      PdfWriter.getInstance(pdfDoc, new FileOutputStream(outputPath));
      pdfDoc.open();

      if (inputPath.toLowerCase().endsWith(".docx")) {
        convertDocxToPdf(inputPath, pdfDoc);
      } else if (inputPath.toLowerCase().endsWith(".doc")) {
        convertDocToPdf(inputPath, pdfDoc);
      }

      pdfDoc.close();
      System.out.println("Conversion completed with images!");

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void convertDocxToPdf(String inputPath, Document pdfDoc) throws Exception {
    FileInputStream fis = new FileInputStream(inputPath);
    XWPFDocument document = new XWPFDocument(fis);

    // Process paragraphs
    for (XWPFParagraph paragraph : document.getParagraphs()) {
      String text = paragraph.getText();
      if (text != null && !text.trim().isEmpty()) {
        pdfDoc.add(new Paragraph(text));
      }

      // Check for images in paragraph
      for (XWPFRun run : paragraph.getRuns()) {
        List<XWPFPicture> pictures = run.getEmbeddedPictures();
        for (XWPFPicture picture : pictures) {
          byte[] imageData = picture.getPictureData().getData();
          try {
            Image image = Image.getInstance(imageData);
            // Scale image if too large
            if (image.getWidth() > 500) {
              image.scaleToFit(500, 400);
            }
            pdfDoc.add(image);
          } catch (Exception e) {
            System.out.println("Could not add image: " + e.getMessage());
          }
        }
      }
    }

    // Process tables
    for (XWPFTable table : document.getTables()) {
      processTable(table, pdfDoc);
    }

    document.close();
    fis.close();
  }

  private static void convertDocToPdf(String inputPath, Document pdfDoc) throws Exception {
    FileInputStream fis = new FileInputStream(inputPath);
    HWPFDocument document = new HWPFDocument(fis);

    // Extract text
    String text = document.getDocumentText();
    if (text != null && !text.trim().isEmpty()) {
      pdfDoc.add(new Paragraph(text));
    }

    // Extract pictures
    List<Picture> pictures = document.getPicturesTable().getAllPictures();
    for (Picture picture : pictures) {
      try {
        byte[] imageData = picture.getContent();
        Image image = Image.getInstance(imageData);
        if (image.getWidth() > 500) {
          image.scaleToFit(500, 400);
        }
        pdfDoc.add(image);
      } catch (Exception e) {
        System.out.println("Could not add image: " + e.getMessage());
      }
    }

    document.close();
    fis.close();
  }

  private static void processTable(XWPFTable table, Document pdfDoc) throws DocumentException {
    PdfPTable pdfTable =
        new com.itextpdf.text.pdf.PdfPTable(table.getRow(0).getTableCells().size());

    for (XWPFTableRow row : table.getRows()) {
      for (XWPFTableCell cell : row.getTableCells()) {
        pdfTable.addCell(cell.getText());
      }
    }
    pdfDoc.add(pdfTable);
  }
}
