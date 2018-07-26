package kz.madibek;

import com.google.zxing.*;
import com.google.zxing.client.j2se.BufferedImageLuminanceSource;
import com.google.zxing.common.HybridBinarizer;
import io.nayuki.qrcodegen.QrCode;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

/**
 * Hello world!
 *
 */
public class App 
{

    public static final String IMG_FILE_NAME = "qr-code-demo.png";
    public static final String IMG_FILE_1 = "C:\\Users\\Public\\Pictures\\Sample Pictures\\1.png";

    public static void main(String[] args ) throws IOException, NotFoundException, InvalidFormatException {

        System.out.println( "QR Code Demo!" );

        String textToQRCode = "https://madibek.github.io/my-site";
        generateQRCode(textToQRCode, IMG_FILE_NAME);
        System.out.println("\nQR Code generated successfully!");
        System.out.println("\nQR Code generated successfully!");

        String text = decodeQRCode(IMG_FILE_NAME);
        System.out.println("\nDecoded QR code text: " + text);

        File image = new File(IMG_FILE_1);
        WordDocument.addImagesToDocument(image, image);
    }

    private static boolean generateQRCode(String text, String imgFileName) throws IOException {

        QrCode demo = QrCode.encodeText(text, QrCode.Ecc.MEDIUM);
        BufferedImage img = demo.toImage(4, 4);
        String fileExtension = getImageExtension(imgFileName);

        return ImageIO.write(img, fileExtension, new File(imgFileName));
    }

    private static String getImageExtension(String imgFileName) {

        String extension;
        if (imgFileName.endsWith(".emf"))
            extension = "emf";
        else if (imgFileName.endsWith(".wmf"))
            extension = "wmf";
        else if (imgFileName.endsWith(".pict"))
            extension = "pict";
        else if (imgFileName.endsWith(".jpeg"))
            extension = "jpeg";
        else if (imgFileName.endsWith(".jpg"))
            extension = "jpg";
        else if (imgFileName.endsWith(".png"))
            extension = "png";
        else if (imgFileName.endsWith(".dib"))
            extension = "dib";
        else if (imgFileName.endsWith(".gif"))
            extension = "gif";
        else if (imgFileName.endsWith(".tiff"))
            extension = "tiff";
        else if (imgFileName.endsWith(".eps"))
            extension = "eps";
        else if (imgFileName.endsWith(".bmp"))
            extension = "bmp";
        else if (imgFileName.endsWith(".wpg"))
            extension = "wpg";
        else {
            throw new RuntimeException("Specify file extension");
        }
        return extension;
    }

    private static String decodeQRCode(String fileName) throws IOException, NotFoundException {

        BufferedImage bufferedImg = ImageIO.read(new File(fileName));
        LuminanceSource source = new BufferedImageLuminanceSource(bufferedImg);
        BinaryBitmap bitmap = new BinaryBitmap(new HybridBinarizer(source));
        Result result = new MultiFormatReader().decode(bitmap);

        return result.getText();
    }
}
