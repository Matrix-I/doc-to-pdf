import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import org.apache.poi.xwpf.usermodel.*;

public class PoiITextDocxToPdfConverter {
  public static void main(String[] args) throws Exception {
    try (FileInputStream fis = new FileInputStream("file-sample.docx");
        XWPFDocument docx = new XWPFDocument(fis);
        FileOutputStream fos = new FileOutputStream("output_poi_itext.pdf");
        PdfWriter writer = new PdfWriter(fos);
        PdfDocument pdf = new PdfDocument(writer);
        Document document = new Document(pdf)) {
      List<IBodyElement> bodyElements = docx.getBodyElements();

      for (IBodyElement element : bodyElements) {
        if (element instanceof XWPFParagraph paragraph) {
          // Handle images within paragraphs
          for (XWPFRun run : paragraph.getRuns()) {
            for (XWPFPicture picture : run.getEmbeddedPictures()) {
              byte[] pictureData = picture.getPictureData().getData();
              Image image = new Image(ImageDataFactory.create(pictureData));
              document.add(image);
            }
          }
          document.add(new Paragraph(paragraph.getText()));
        } else if (element instanceof XWPFTable xwpfTable) {
          Table pdfTable = new Table(xwpfTable.getRow(0).getTableCells().size());

          for (XWPFTableRow xwpfRow : xwpfTable.getRows()) {
            for (XWPFTableCell xwpfCell : xwpfRow.getTableCells()) {
              pdfTable.addCell(new Paragraph(xwpfCell.getText()));
            }
          }
          document.add(pdfTable);
        }
      }
      System.out.println("Conversion to PDF completed successfully using Apache POI and iText!");
    }
  }
}
