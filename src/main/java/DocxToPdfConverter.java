import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.*;
import org.apache.poi.xwpf.usermodel.*;

public class DocxToPdfConverter {

  public static void main(String[] args) {
    String inputPath = "file-sample.docx";
    String outputPath = "output222.pdf";

    try (InputStream inputStream = new FileInputStream(inputPath);
        OutputStream outputStream = new FileOutputStream(outputPath)) {

      // Load DOCX file into XWPFDocument
      XWPFDocument document = new XWPFDocument(inputStream);

      // Create PDF conversion options
      PdfOptions options = PdfOptions.create();

      // Convert DOCX to PDF
      PdfConverter.getInstance().convert(document, outputStream, options);

      System.out.println("Conversion completed successfully: " + outputPath);

    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
