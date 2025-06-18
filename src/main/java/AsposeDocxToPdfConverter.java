import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

public class AsposeDocxToPdfConverter {
  public static void main(String[] args) throws Exception {
    // Load the input DOCX file
    Document doc = new Document("file-sample.docx");

    // Save the document as PDF
    doc.save("output_aspose.pdf", SaveFormat.PDF);

    System.out.println("Conversion to PDF completed successfully using Aspose.Words!");
  }
}
