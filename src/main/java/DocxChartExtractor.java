import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class DocxChartExtractor {

  public static List<XWPFChart> extractCharts(String docxFilePath) throws IOException {
    try (FileInputStream fis = new FileInputStream(docxFilePath);
        XWPFDocument document = new XWPFDocument(fis)) {
      return document.getCharts();
    }
  }

  public static void main(String[] args) {
    try {
      List<XWPFChart> charts = extractCharts("file-sample.docx");
      if (!charts.isEmpty()) {
        System.out.println(charts.get(0));
        System.out.println("Found " + charts.size() + " chart(s) in the document.");
        // You can now process each XWPFChart object
      } else {
        System.out.println("No charts found in the document.");
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
