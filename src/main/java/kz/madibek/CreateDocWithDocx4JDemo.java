package kz.madibek;


import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.*;
import java.util.List;

import static kz.madibek.App.IMG_FILE_NAME;

public class CreateDocWithDocx4JDemo {

    private static ObjectFactory objectFactory = new ObjectFactory();

    /**
     * As is usual, we create a package to contain the document.
     * Then we create a file that contains the image we want to add to the document.
     * In order to be able to do something with this image, we'll have to convert
     * it to an array of bytes. Finally we add the image to the package
     * and save the package.
     */
    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage wordMLPackage =
                WordprocessingMLPackage.createPackage();

        File file = new File(IMG_FILE_NAME);
        byte[] bytes = convertImageToByteArray(file);
        // 1. the Header part
        Relationship relationship = createFooterPart(wordMLPackage, bytes);
        // 2. an entry in SectPr
        createFooterReference(wordMLPackage, relationship);

        wordMLPackage.save(new java.io.File("1.docx"));
    }

    /**
     * Docx4j contains a utility method to create an image part from an array of
     * bytes and then adds it to the given package. In order to be able to add this
     * image to a paragraph, we have to convert it into an inline object. For this
     * there is also a method, which takes a filename hint, an alt-text, two ids
     * and an indication on whether it should be embedded or linked to.
     * One id is for the drawing object non-visual properties of the document, and
     * the second id is for the non visual drawing properties of the picture itself.
     * Finally we add this inline object to the paragraph and the paragraph to the
     * main document of the package.
     *
     * @param wordMLPackage The package we want to add the image to
     * @param bytes         The bytes of the image
     * @throws Exception Sadly the createImageInline method throws an Exception
     *                   (and not a more specific exception type)
     */
    private static void addImageToPackage(WordprocessingMLPackage wordMLPackage,
                                          byte[] bytes) throws Exception {
        BinaryPartAbstractImage imagePart =
                BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        int docPrId = 1;
        int cNvPrId = 2;
        Inline inline = imagePart.createImageInline("Filename hint",
                "Alternative text", docPrId, cNvPrId, false);

        P paragraph = addInlineImageToParagraph(inline);

        wordMLPackage.getMainDocumentPart().addObject(paragraph);
    }

    /**
     * We create an object factory and use it to create a paragraph and a run.
     * Then we add the run to the paragraph. Next we create a drawing and
     * add it to the run. Finally we add the inline object to the drawing and
     * return the paragraph.
     *
     * @param inline The inline object containing the image.
     * @return the paragraph containing the image
     */
    private static P addInlineImageToParagraph(Inline inline) {
        // Now add the in-line image to a paragraph
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();
        R run = factory.createR();
        paragraph.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return paragraph;
    }

    /**
     * Convert the image from the file into an array of bytes.
     *
     * @param file the image file to be converted
     * @return the byte array containing the bytes from the image
     * @throws FileNotFoundException
     * @throws IOException
     */
    private static byte[] convertImageToByteArray(File file)
            throws FileNotFoundException, IOException {
        InputStream is = new FileInputStream(file);
        long length = file.length();
        // You cannot create an array using a long, it needs to be an int.
        if (length > Integer.MAX_VALUE) {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int) length];
        int offset = 0;
        int numRead = 0;
        while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0) {
            offset += numRead;
        }
        // Ensure all the bytes have been read
        if (offset < bytes.length) {
            System.out.println("Could not completely read file "
                    + file.getName());
        }
        is.close();
        return bytes;
    }

    public static Relationship createFooterPart(
            WordprocessingMLPackage wordprocessingMLPackage, byte[] bytesImage)
            throws Exception {

        FooterPart footerPart = new FooterPart();
        Relationship rel = wordprocessingMLPackage.getMainDocumentPart()
                .addTargetPart(footerPart);

        // After addTargetPart, so image can be added properly
        footerPart.setJaxbElement(getFtr(wordprocessingMLPackage, footerPart, bytesImage));

        return rel;
    }

    public static void createFooterReference(
            WordprocessingMLPackage wordprocessingMLPackage,
            Relationship relationship)
            throws InvalidFormatException {

        List<SectionWrapper> sections = wordprocessingMLPackage.getDocumentModel().getSections();

        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        // There is always a section wrapper, but it might not contain a sectPr
        if (sectPr == null) {
            sectPr = objectFactory.createSectPr();
            wordprocessingMLPackage.getMainDocumentPart().addObject(sectPr);
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }

        FooterReference footerReference = objectFactory.createFooterReference();
        footerReference.setId(relationship.getId());
        footerReference.setType(HdrFtrRef.DEFAULT);
        sectPr.getEGHdrFtrReferences().add(footerReference);// add header or
        // footer references
    }


    public static Ftr getFtr(WordprocessingMLPackage wordprocessingMLPackage,
                             Part sourcePart, byte[] bytesImage) throws Exception {

        Ftr ftr = objectFactory.createFtr();

        ftr.getContent().add(
                newImage(wordprocessingMLPackage,
                        sourcePart,
//                        BufferUtil.getBytesFromInputStream(is),
                        bytesImage,
                        "filename", "alttext", 1, 2
                )
        );
        return ftr;
    }

    public static org.docx4j.wml.P newImage(WordprocessingMLPackage wordMLPackage,
                                            Part sourcePart,
                                            byte[] bytes,
                                            String filenameHint, String altText,
                                            int id1, int id2) throws Exception {

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage,
                sourcePart, bytes);

        Inline inline = imagePart.createImageInline(filenameHint, altText,
                id1, id2, false);

        // Now add the inline in w:p/w:r/w:drawing
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return p;
    }
}