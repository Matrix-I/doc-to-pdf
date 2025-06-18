import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class SmartDocConverter {

  public static void convertToPdf(String inputPath, String outputPath) {
    try {
      // Use BufferedInputStream which supports mark()
      FileInputStream fis = new FileInputStream(inputPath);
      BufferedInputStream bis = new BufferedInputStream(fis);

      FileMagic fileMagic = FileMagic.valueOf(bis);
      bis.close();
      fis.close();

      String text = "";

      switch (fileMagic) {
        case OLE2:
          // It's a .doc file
          text = extractFromDoc(inputPath);
          break;
        case OOXML:
          // It's a .docx file
          text = extractFromDocx(inputPath);
          break;
        default:
          throw new IllegalArgumentException("Unsupported file type: " + fileMagic);
      }

      createPdf(text, outputPath);
      System.out.println("Conversion completed successfully!");

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static String extractFromDoc(String filePath) throws Exception {
    FileInputStream fis = new FileInputStream(filePath);
    HWPFDocument document = new HWPFDocument(fis);
    WordExtractor extractor = new WordExtractor(document);
    String text = extractor.getText();

    extractor.close();
    document.close();
    fis.close();

    return text;
  }

  private static String extractFromDocx(String filePath) throws Exception {
    FileInputStream fis = new FileInputStream(filePath);
    XWPFDocument document = new XWPFDocument(fis);
    XWPFWordExtractor extractor = new XWPFWordExtractor(document);
    String text = extractor.getText();

    extractor.close();
    document.close();
    fis.close();

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
