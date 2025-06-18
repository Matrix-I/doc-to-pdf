import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class UniversalDocToPdfConverter {

  public static void convertToPdf(String inputPath, String outputPath) {
    try {
      String text = extractText(inputPath);
      createPdf(text, outputPath);
      System.out.println("Conversion completed successfully!");
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static String extractText(String filePath) throws Exception {
    FileInputStream fis = new FileInputStream(filePath);
    String text;

    try {
      if (filePath.toLowerCase().endsWith(".docx")) {
        // Handle DOCX files
        XWPFDocument document = new XWPFDocument(fis);
        XWPFWordExtractor extractor = new XWPFWordExtractor(document);
        text = extractor.getText();
        extractor.close();
        document.close();
      } else if (filePath.toLowerCase().endsWith(".doc")) {
        // Handle DOC files
        HWPFDocument document = new HWPFDocument(fis);
        WordExtractor extractor = new WordExtractor(document);
        text = extractor.getText();
        extractor.close();
        document.close();
      } else {
        throw new IllegalArgumentException("Unsupported file format. Use .doc or .docx files.");
      }
    } finally {
      fis.close();
    }

    return text;
  }

  private static void createPdf(String text, String outputPath) throws Exception {
    Document pdfDoc = new Document();
    PdfWriter.getInstance(pdfDoc, new FileOutputStream(outputPath));
    pdfDoc.open();
    pdfDoc.add(new Paragraph(text));
    pdfDoc.close();
  }
}
