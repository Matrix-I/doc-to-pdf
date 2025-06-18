import java.nio.file.Paths;

public class DockerLibreOfficeConverter {

  public static void main(String[] args) {
    convertToPdf(
        "/Users/linh.nguyen/Desktop/personal/doc-to-pdf/file-sample.docx",
        "/Users/linh.nguyen/Desktop/personal/doc-to-pdf");
  }

  public static void convertToPdf(String inputPath, String outputPath) {
    try {
      String inputDir = Paths.get(inputPath).getParent().toString();
      String inputFileName = Paths.get(inputPath).getFileName().toString();
      String outputDir = Paths.get(outputPath).toString();

      ProcessBuilder pb =
          new ProcessBuilder(
              "docker",
              "run",
              "--rm",
              "-v",
              inputDir + ":/input",
              "-v",
              outputDir + ":/output",
              "linuxserver/libreoffice",
              "libreoffice",
              "--headless",
              "--convert-to",
              "pdf",
              "--outdir",
              "/output",
              "/input/" + inputFileName);

      Process process = pb.start();
      process.waitFor();

    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}
