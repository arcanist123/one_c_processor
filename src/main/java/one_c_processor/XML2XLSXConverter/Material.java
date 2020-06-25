package one_c_processor.XML2XLSXConverter;

import one_c_processor.en_tag;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;


import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.Vector;

public class Material {
    private static final short MATERIAL_NAME_COLUMN = 5;
    private static final short MATERIAL_VENDOR_CODE_COLUMN = 4;
    private static final double PICTURE_HEIGHT = 100;
    byte[] picture;
    private String materialName;
    private String materialGUID;
    private String materialVendorCode;
    private Vector<MaterialAttribute> materialAttributes;

    private Material() {
        this.materialAttributes = new Vector<MaterialAttribute>(0);
    }

    public static Material createMaterial(Element materialElement) throws IOException {

        //create instance of material
        Material material = new Material();

        //get the name of material
        material.materialName = materialElement.getElementsByTagName(en_tag.MATERIAL_NAME.toString()).item(0).getTextContent();

        //get the guid of material
        material.materialGUID = materialElement.getElementsByTagName(en_tag.MATERIAL_GUID.toString()).item(0).getTextContent();

        //get the material vendor code
        material.materialVendorCode = materialElement.getElementsByTagName(en_tag.MATERIAL_VENDOR_CODE.toString()).item(0).getTextContent();

        //get the picture
        //get the possible pictures
        NodeList pictures = materialElement.getElementsByTagName(en_tag.PICTURE.toString());
        //check if the current material has picture
        if (pictures.getLength() > 0) {
            //this material has picture. get it
            String pictureEncoded = pictures.item(0).getTextContent();
            //decode base64 string
            material.picture = Base64.getDecoder().decode(pictureEncoded);
            //ensure that the size of the picture is adequate
            material.picture = resizePicture(material.picture);

        } else {
            ///no picture was provided from source
        }

        //now populate the list of the material attributes
        NodeList materialAttributes = materialElement.getElementsByTagName(en_tag.ATTRIBUTE.toString());

        for (int i = 0; i < materialAttributes.getLength(); i++) {
            //get the current material attribute
            Element materialAttributeElement = (Element) materialAttributes.item(i);
            //create an instance of material attribute
            material.materialAttributes.add(MaterialAttribute.createMaterialAttribute(materialAttributeElement));
        }

        //return the results of processing
        return material;
    }

    private static byte[] resizePicture(byte[] picture) throws IOException {
        //check if the image is too big
        if (picture.length > 50000) {

            int factor = 1;
            double length = picture.length;
            byte[] currentPicture = picture;
            while (length > 50000) {
                factor++;
                currentPicture = resizePictureByFactor(picture, factor);
                length = currentPicture.length;
                if (factor > 10) {
                    return picture;
                }
            }

            return currentPicture;

        } else {
            return picture;
        }
    }

    private static byte[] resizePictureByFactor(byte[] picture, int factor) throws IOException {

        BufferedImage inputImage = createImageFromBytes(picture);

        int scaledWidth = inputImage.getWidth() / factor;
        int scaledHeight = inputImage.getHeight() / factor;
        BufferedImage outputImage = new BufferedImage(scaledWidth, scaledHeight, inputImage.getType());
        Graphics2D g2d = outputImage.createGraphics();
        g2d.drawImage(inputImage, 0, 0, scaledWidth, scaledHeight, null);
        g2d.dispose();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(outputImage, "png", baos);
        baos.flush();
        byte[] imageInByte = baos.toByteArray();
        baos.close();

        return imageInByte;
    }

    private static BufferedImage createImageFromBytes(byte[] imageData) {
        ByteArrayInputStream bais = new ByteArrayInputStream(imageData);
        try {
            return ImageIO.read(bais);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public short writeToXLSX(Sheet mainSheet, short currentRow, cl_styles stylesCache) {

        //materials themselves are not being written to the sheet - only the underlying attributes are
        //save the current row for future usage
        short originalRow = currentRow;

        //write the attributes of material
        for (int i = 0; i < this.materialAttributes.size(); i++) {
            //write the current material attribute
            currentRow = this.materialAttributes.elementAt(i).writeToXLSX(mainSheet, currentRow, stylesCache);
        }

        //determine the height of the row
        float rowHeith;
        if (materialAttributes.size() > 5) {
            //set the standard row height
            rowHeith = (float) PICTURE_HEIGHT / 5;
        } else {
            rowHeith = (float) PICTURE_HEIGHT / materialAttributes.size();
        }

        //output the material guid and material name
        for (int i = originalRow; i < currentRow; i++) {
            //get the current row
            Row row = mainSheet.getRow(i);

            //set the row height
            row.setHeightInPoints(rowHeith);
            //output the material guid
            Cell materialGuidCell = row.createCell(ColumnIndex.HEADER_MATERIAL_GUID);
            materialGuidCell.setCellValue(this.materialGUID);

            //output the material name
            Cell materialNameCell = row.createCell(MATERIAL_NAME_COLUMN);
            materialNameCell.setCellValue(this.materialName);
            materialNameCell.setCellStyle(stylesCache.getMaterialNameCellStyle());

            //output the vendor code
            Cell materialVendorCodeCell = row.createCell(MATERIAL_VENDOR_CODE_COLUMN);
            materialVendorCodeCell.setCellValue(this.materialVendorCode);
            CellStyle materialVendorCodeCellStyle = stylesCache.getMaterialVendorCodeCellStyle();
            materialVendorCodeCell.setCellStyle(materialVendorCodeCellStyle);
        }

        //unite the cells for material
        if (originalRow == currentRow - 1) {
            //in this case we are not having more then one row to work. no need to unite one cell
        } else {
            //we have some cell to merge. merge them
            //merge the material vendor code
            CellRangeAddress materialVendorCodeCellRange = new CellRangeAddress(originalRow, currentRow - 1, ColumnIndex.HEADER_VENDOR_CODE, ColumnIndex.HEADER_VENDOR_CODE);
            mainSheet.addMergedRegion(materialVendorCodeCellRange);

            //merge the material name cells
            CellRangeAddress materialNameCellRange = new CellRangeAddress(originalRow, currentRow - 1, ColumnIndex.HEADER_MATERIAL_NAME, ColumnIndex.HEADER_MATERIAL_NAME);
            mainSheet.addMergedRegion(materialNameCellRange);

        }
        if (this.picture == null) {

        } else {
            this.insertPicture(mainSheet, this.picture, originalRow, currentRow);
        }
        return currentRow;
    }

    private void insertPicture(Sheet mainSheet, byte[] picture, short originalRow, short currentRow) {
        //add the picture to the workbook
        //get picture cache instance
        PictureCache cache = PictureCache.getInstance();

        int pictureId = cache.addPicture(mainSheet.getWorkbook(), picture, Workbook.PICTURE_TYPE_JPEG);

        CreationHelper helper = mainSheet.getWorkbook().getCreationHelper();
        Drawing drawing = mainSheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(ColumnIndex.HEADER_PHOTO);
        anchor.setCol2(ColumnIndex.HEADER_PHOTO + 1);
        anchor.setRow1(originalRow);
        anchor.setRow2(currentRow);

        Picture pict = drawing.createPicture(anchor, pictureId);
    }
}
