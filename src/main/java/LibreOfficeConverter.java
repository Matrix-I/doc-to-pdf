import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

public class LibreOfficeConverter {

  public static boolean convertToPdf(String inputPath, String outputDir) {
    try {
      // Ensure output directory exists
      Files.createDirectories(Paths.get(outputDir));

      ProcessBuilder pb =
          new ProcessBuilder(
              "soffice", // or "libreoffice" on some systems
              "--headless",
              "--convert-to",
              "pdf",
              "--outdir",
              outputDir,
              inputPath);

      pb.redirectErrorStream(true);
      Process process = pb.start();

      // Read output for debugging
      BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
      String line;
      while ((line = reader.readLine()) != null) {
        System.out.println(line);
      }

      int exitCode = process.waitFor();

      if (exitCode == 0) {
        System.out.println("Conversion successful - images and charts preserved!");
        return true;
      } else {
        System.out.println("Conversion failed with exit code: " + exitCode);
        return false;
      }

    } catch (Exception e) {
      e.printStackTrace();
      return false;
    }
  }

  public static void main(String[] args) {
    String inputFile = "file-sample.docx";
    String outputDir = "/Users/linh.nguyen/Desktop/personal/doc-to-pdf";

    convertToPdf(inputFile, outputDir);
  }
}
