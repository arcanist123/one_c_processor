package one_c_processor.XML2XLSXConverter;

import one_c_processor.en_tag;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;

import java.io.IOException;

public class PriceList {
    private DocumentHeader documentHeader;
    private MaterialsSection materialsSection;

    public PriceList(Element document) throws IOException {
        //ensure that the document is correctly formatted
        document.normalize();

        //get the header
        Element documentHeader = (Element) document.getElementsByTagName(en_tag.ORGANISATION.toString()).item(0);
        this.documentHeader = new DocumentHeader(documentHeader);

        //get the materialsSection
        Element materials = (Element) document.getElementsByTagName(en_tag.MATERIALS_SECTION.toString()).item(0);
        this.materialsSection = MaterialsSection.createMaterials(materials);

    }

    public PriceList(Sheet worksheet){

    }

    public void writeToXLSX(Sheet mainSheet) {

        //set the sizes of the columns
        this.setSizes(mainSheet);

        //we are starting to write to the sheet from the first row
        short currentRow = 1;

        //get the instance of the styles cache
        cl_styles stylesCache = cl_styles.createInstance(mainSheet.getWorkbook());

        //output document header
        currentRow = documentHeader.writeToXLSX(mainSheet, currentRow, stylesCache);

        //output materialsSection
        materialsSection.writeToXLSX(mainSheet, currentRow, stylesCache);

    }

    private void setSizes(Sheet mainSheet) {
        //set guid columns hidden
        mainSheet.setColumnHidden(ColumnIndex.HEADER_MATERIAL_GROUP_GUID, true);
        mainSheet.setColumnHidden(ColumnIndex.HEADER_MATERIAL_GUID, true);
        mainSheet.setColumnHidden(ColumnIndex.HEADER_ATTRIBUTE_GUID, true);


        //set the photo to be quadratic
        short photoWidth = PixelUtil.pixel2WidthUnits(150);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_PHOTO, photoWidth);

        //set the width of the vendor code
        short vendorCodeWidth = PixelUtil.pixel2WidthUnits(150);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_VENDOR_CODE, vendorCodeWidth);

        //set the width of material name
        short materialNameWidth = PixelUtil.pixel2WidthUnits(200);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_MATERIAL_NAME, materialNameWidth);

        //set the width of attribute name
        short attributeNameWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_ATTRIBUTE_NAME, attributeNameWidth);

        //set the width of size in santimeter
        short sizeWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_SANTIMETERS, sizeWidth);

        //set the width of price
        short priceWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_PRICE, priceWidth);

        //set the width of the discount
        short discountWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_DISCOUNT, discountWidth);

        //set the width of price with discount
        short priceWithDiscountWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_PRICE_WITH_DISCOUNT, priceWithDiscountWidth);

        //set the width of quantity column
        short quantityWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_QUANTITY, quantityWidth);

        //set the width of order column
        short orderColumnWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_SOLD_QUANTITY, orderColumnWidth);

        //set the width of sum column
        short sumColumnWidth = PixelUtil.pixel2WidthUnits(100);
        mainSheet.setColumnWidth(ColumnIndex.HEADER_SUM, sumColumnWidth);

    }

}