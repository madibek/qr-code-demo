package kz.madibek;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WordDocument {

    public static void addImagesToDocument(File imgFile1, File imgFile2) throws IOException, InvalidFormatException {

        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        String p1 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";

        String p2 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";

        run.setText(p1);
        run.addBreak();
        addImageToRun(run, imgFile1);

        run.addBreak();

        run.setText(p2);
        run.addBreak();
        addImageToRun(run, imgFile2);

        addImageToFooter(document, imgFile2);

        FileOutputStream out = new FileOutputStream("word-qr-code.docx");
        document.write(out);
        out.close();
        document.close();

    }

    private static void addImageToFooter(XWPFDocument document, File imgFile2) throws InvalidFormatException, IOException {

        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);

        // create footer
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);

//        CTTabStop tabStop = paragraph.getCTP().getPPr().addNewTabs().addNewTab();
//        tabStop.setVal(STTabJc.RIGHT);
//        int twipsPerInch =  1440;
//        tabStop.setPos(BigInteger.valueOf(6 * twipsPerInch));
//
//        run = paragraph.createRun();
//        run.setText("QR Code");
//        run.addTab();

        XWPFRun run = paragraph.createRun();
        addImageToRun(run, imgFile2);

    }

    private static void addImageToRun(XWPFRun run, File imgFile) throws IOException, InvalidFormatException {

        BufferedImage image = ImageIO.read(imgFile);
        int width = image.getWidth();
        int height = image.getHeight();
        String imgName = imgFile.getName();
        int formatImg = getImageFormat(imgName);

        run.addPicture(new FileInputStream(imgFile), formatImg, imgName, Units.toEMU(width), Units.toEMU(height));

    }

    private static int getImageFormat(String imgFileName) {

        int format;
        if (imgFileName.endsWith(".emf"))
            format = XWPFDocument.PICTURE_TYPE_EMF;
        else if (imgFileName.endsWith(".wmf"))
            format = XWPFDocument.PICTURE_TYPE_WMF;
        else if (imgFileName.endsWith(".pict"))
            format = XWPFDocument.PICTURE_TYPE_PICT;
        else if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg"))
            format = XWPFDocument.PICTURE_TYPE_JPEG;
        else if (imgFileName.endsWith(".png"))
            format = XWPFDocument.PICTURE_TYPE_PNG;
        else if (imgFileName.endsWith(".dib"))
            format = XWPFDocument.PICTURE_TYPE_DIB;
        else if (imgFileName.endsWith(".gif"))
            format = XWPFDocument.PICTURE_TYPE_GIF;
        else if (imgFileName.endsWith(".tiff"))
            format = XWPFDocument.PICTURE_TYPE_TIFF;
        else if (imgFileName.endsWith(".eps"))
            format = XWPFDocument.PICTURE_TYPE_EPS;
        else if (imgFileName.endsWith(".bmp"))
            format = XWPFDocument.PICTURE_TYPE_BMP;
        else if (imgFileName.endsWith(".wpg"))
            format = XWPFDocument.PICTURE_TYPE_WPG;
        else {
            return 0;
        }
        return format;
    }

}
