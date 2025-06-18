import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

public class SimplifiedDocxToPdfConverter {

  public static void main(String[] args) {
    try {
      String docxPath = "file-sample.docx";
      String pdfPath = "output222.pdf";

      convertDocxImagesToPdf(docxPath, pdfPath);
      System.out.println("Conversion completed successfully!");

    } catch (Exception e) {
      System.err.println("Error during conversion: " + e.getMessage());
      e.printStackTrace();
    }
  }

  public static void convertDocxImagesToPdf(String docxPath, String pdfPath) throws Exception {
    // Extract images from DOCX
    List<byte[]> imageData = extractAllImagesFromDocx(docxPath);

    // Create PDF with extracted images
    createPdfWithExtractedImages(imageData, pdfPath);
  }

  private static List<byte[]> extractAllImagesFromDocx(String docxPath) throws Exception {
    List<byte[]> imageDataList = new ArrayList<>();

    try (FileInputStream fis = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fis)) {

      System.out.println("Successfully opened DOCX document");

      // Extract all embedded pictures (including charts saved as images)
      List<XWPFPictureData> pictures = document.getAllPictures();
      System.out.println("Found " + pictures.size() + " pictures in document");

      for (int i = 0; i < pictures.size(); i++) {
        XWPFPictureData picture = pictures.get(i);
        byte[] data = picture.getData();

        if (data != null && data.length > 0) {
          System.out.println(
              "Extracted image "
                  + (i + 1)
                  + " - Type: "
                  + picture.getPictureType()
                  + " - Size: "
                  + data.length
                  + " bytes");
          imageDataList.add(data);
        }
      }

      // If no images found, create a placeholder
      if (imageDataList.isEmpty()) {
        System.out.println("No images found, creating placeholder");
        BufferedImage placeholder = createNoImagesPlaceholder();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(placeholder, "PNG", baos);
        imageDataList.add(baos.toByteArray());
      }
    }

    return imageDataList;
  }

  private static BufferedImage createNoImagesPlaceholder() {
    int width = 600;
    int height = 400;

    BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
    Graphics2D g2d = image.createGraphics();

    // Set rendering hints
    g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
    g2d.setRenderingHint(
        RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

    // Fill background
    g2d.setColor(Color.WHITE);
    g2d.fillRect(0, 0, width, height);

    // Draw border
    g2d.setColor(Color.LIGHT_GRAY);
    g2d.drawRect(20, 20, width - 40, height - 40);

    // Draw message
    g2d.setColor(Color.DARK_GRAY);
    g2d.setFont(new Font("Arial", Font.BOLD, 24));
    String message = "No Images Found";
    int messageWidth = g2d.getFontMetrics().stringWidth(message);
    g2d.drawString(message, (width - messageWidth) / 2, height / 2 - 20);

    g2d.setFont(new Font("Arial", Font.PLAIN, 16));
    String subtitle = "in DOCX Document";
    int subtitleWidth = g2d.getFontMetrics().stringWidth(subtitle);
    g2d.drawString(subtitle, (width - subtitleWidth) / 2, height / 2 + 10);

    g2d.dispose();
    return image;
  }

  private static void createPdfWithExtractedImages(List<byte[]> imageDataList, String pdfPath)
      throws Exception {
    PdfWriter writer = new PdfWriter(pdfPath);
    PdfDocument pdfDoc = new PdfDocument(writer);
    Document document = new Document(pdfDoc);

    // Add title
    document.add(
        new Paragraph("Images Extracted from DOCX Document")
            .setBold()
            .setFontSize(18)
            .setMarginBottom(20));

    // Add each image to the PDF
    for (int i = 0; i < imageDataList.size(); i++) {
      byte[] imageBytes = imageDataList.get(i);

      try {
        // Create iText Image
        ImageData imageData = ImageDataFactory.create(imageBytes);
        Image image = new Image(imageData);

        // Calculate scaling to fit page
        float pageWidth = document.getPdfDocument().getDefaultPageSize().getWidth() - 80;
        float pageHeight = document.getPdfDocument().getDefaultPageSize().getHeight() - 200;

        float imageWidth = image.getImageWidth();
        float imageHeight = image.getImageHeight();

        // Scale image if it's too large
        if (imageWidth > pageWidth || imageHeight > pageHeight) {
          float scaleX = pageWidth / imageWidth;
          float scaleY = pageHeight / imageHeight;
          float scale = Math.min(scaleX, scaleY);

          image.scale(scale, scale);
        }

        // Add image title
        document.add(
            new Paragraph("Image " + (i + 1))
                .setBold()
                .setFontSize(14)
                .setMarginTop(10)
                .setMarginBottom(5));

        // Add dimensions info
        document.add(
            new Paragraph(
                    String.format("Original size: %.0f x %.0f pixels", imageWidth, imageHeight))
                .setFontSize(10)
                .setMarginBottom(10));

        // Add image to document
        document.add(image);

        // Add page break if not the last image
        if (i < imageDataList.size() - 1) {
          document.add(new com.itextpdf.layout.element.AreaBreak());
        }

      } catch (Exception e) {
        // If image can't be processed, add error message
        document.add(
            new Paragraph("Error processing image " + (i + 1) + ": " + e.getMessage())
                .setItalic()
                .setFontColor(com.itextpdf.kernel.colors.ColorConstants.RED));
        System.err.println("Failed to process image " + (i + 1) + ": " + e.getMessage());
      }
    }

    document.close();
    System.out.println("PDF created successfully: " + pdfPath);
  }

  // Alternative method to extract specific file types
  public static void extractSpecificImageTypes(String docxPath, String outputDir) throws Exception {
    try (FileInputStream fis = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fis)) {

      List<XWPFPictureData> pictures = document.getAllPictures();

      for (int i = 0; i < pictures.size(); i++) {
        XWPFPictureData picture = pictures.get(i);
        String extension = getFileExtension(picture.getPictureType());

        if (extension != null) {
          String fileName = outputDir + "/image_" + (i + 1) + extension;

          try (FileOutputStream fos = new FileOutputStream(fileName)) {
            fos.write(picture.getData());
            System.out.println("Saved: " + fileName);
          }
        }
      }
    }
  }

  private static String getFileExtension(int pictureType) {
    switch (pictureType) {
        //            case XWPFDocument.PICTURE_TYPE_PNG:
        //                return ".png";
        //            case XWPFDocument.PICTURE_TYPE_JPEG:
        //                return ".jpg";
        //            case XWPFDocument.PICTURE_TYPE_GIF:
        //                return ".gif";
        //            case XWPFDocument.PICTURE_TYPE_BMP:
        //                return ".bmp";
        //            case XWPFDocument.PICTURE_TYPE_WMF:
        //                return ".wmf";
        //            case XWPFDocument.PICTURE_TYPE_EMF:
        //                return ".emf";
      default:
        return ".jpg";
    }
  }
}
