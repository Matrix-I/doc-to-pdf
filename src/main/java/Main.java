public class Main {
  public static void main(String[] args) {
    String inputPath = "file-sample.docx";
    String outputPath = "output.pdf";
    // This will work for both .doc and .docx files
    //        UniversalDocToPdfConverter.convertToPdf(inputPath, outputPath);
    //    DockerLibreOfficeConverter.convertToPdf(
    //        "/Users/linh.nguyen/Desktop/personal/doc-to-pdf/file-sample.docx",
    //        "/Users/linh.nguyen/Desktop/personal/doc-to-pdf");
    // or
    //    ChartPreservingConverter.convertToPdf(inputPath, outputPath);
  }
}
