import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.data.category.DefaultCategoryDataset;

public class ChartToDocxToPdf {
  public static void main(String[] args) {
    String templatePath = "file-sample.docx";
    String outputPath = "output_with_chart.pdf";

    try (FileInputStream inputStream = new FileInputStream(templatePath);
        XWPFDocument document = new XWPFDocument(inputStream);
        FileOutputStream outputStream = new FileOutputStream(outputPath)) {

      // Create a simple chart using JFreeChart
      DefaultCategoryDataset dataset = new DefaultCategoryDataset();
      dataset.addValue(1.0, "Series1", "Category1");
      dataset.addValue(4.0, "Series1", "Category2");
      dataset.addValue(3.0, "Series1", "Category3");

      JFreeChart chart = ChartFactory.createBarChart("Sample Chart", "Category", "Value", dataset);

      // Convert chart to PNG image
      ByteArrayOutputStream chartOut = new ByteArrayOutputStream();
      ChartUtils.writeChartAsPNG(chartOut, chart, 500, 300);
      byte[] chartBytes = chartOut.toByteArray();

      // Add chart image to DOCX
      XWPFParagraph paragraph = document.createParagraph();
      XWPFRun run = paragraph.createRun();
      run.addPicture(
          new ByteArrayInputStream(chartBytes),
          XWPFDocument.PICTURE_TYPE_PNG,
          "chart.png",
          Units.toEMU(500),
          Units.toEMU(300));

      // Convert DOCX to PDF
      PdfOptions options = PdfOptions.create();
      PdfConverter.getInstance().convert(document, outputStream, options);

      System.out.println("PDF with chart created: " + outputPath);

    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
