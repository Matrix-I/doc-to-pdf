import java.io.File;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;

public class JodConverterExample {

  public static void main(String[] args) throws Exception {
    // Create an OfficeManager instance.
    // This will find a LibreOffice installation automatically.
    LocalOfficeManager officeManager = LocalOfficeManager.builder().build();

    try {
      // Start the office process
      officeManager.start();

      System.out.println("Converting DOCX to PDF using JODConverter...");

      // Create the converter
      LocalConverter converter = LocalConverter.make(officeManager);

      File inputFile = new File("file-sample.docx");
      File outputFile = new File("output_jodconverter.pdf");

      // Convert the file
      converter.convert(inputFile).to(outputFile).execute();

      System.out.println("Conversion successful: " + outputFile.getAbsolutePath());

    } finally {
      // Stop the office process
      if (officeManager.isRunning()) {
        officeManager.stop();
      }
    }
  }
}
