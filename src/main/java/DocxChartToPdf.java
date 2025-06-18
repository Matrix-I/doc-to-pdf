import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class DocxChartToPdf {
  public static void main(String[] args) {
    String docxPath = "file-sample.docx"; // Path to your DOCX file
    String pdfPath = "output123.pdf"; // Path to output PDF file

    try {
      // Step 1: Read the DOCX file
      FileInputStream fis = new FileInputStream(docxPath);
      XWPFDocument doc = new XWPFDocument(fis);

      // Step 2: Create a PDF document
      PdfWriter writer = new PdfWriter(new FileOutputStream(pdfPath));
      PdfDocument pdfDoc = new PdfDocument(writer);
      Document document = new Document(pdfDoc);

      // Step 3: Extract images from DOCX
      List<XWPFParagraph> paragraphs = doc.getParagraphs();
      boolean imageFound = false;

      for (XWPFParagraph paragraph : paragraphs) {
        for (XWPFRun run : paragraph.getRuns()) {
          List<XWPFPicture> pictures = run.getEmbeddedPictures();
          for (XWPFPicture picture : pictures) {
            // Get image data
            byte[] imageBytes = picture.getPictureData().getData();
            String imageType = picture.getPictureData().getFileName().toLowerCase();

            // Step 4: Add image to PDF
            Image image = new Image(ImageDataFactory.create(imageBytes));
            // Optionally scale the image to fit the page
            image.scaleToFit(500, 500);
            document.add(image);
            imageFound = true;
          }
        }
      }

      if (!imageFound) {
        System.out.println("No embedded images found in the DOCX file.");
      }

      // Step 5: Close documents
      document.close();
      pdfDoc.close();
      doc.close();
      fis.close();

      System.out.println("PDF created successfully at: " + pdfPath);

    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
