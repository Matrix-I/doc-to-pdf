import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import java.io.*;

public class Documents4jConverter {

  public static boolean convertToPdf(String inputPath, String outputPath) {
    try {
      File inputFile = new File(inputPath);
      File outputFile = new File(outputPath);

      IConverter converter = LocalConverter.builder().build();

      boolean result =
          converter
              .convert(inputFile)
              .as(DocumentType.DOCX)
              .to(outputFile)
              .as(DocumentType.PDF)
              .execute();

      converter.shutDown();

      if (result) {
        System.out.println("Conversion successful with Documents4j!");
        return true;
      } else {
        System.out.println("Conversion failed");
        return false;
      }

    } catch (Exception e) {
      e.printStackTrace();
      return false;
    }
  }
}
