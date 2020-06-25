package one_c_processor.XML2XLSXConverter;


import one_c_processor.en_tag;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;

public class DocumentHeader {
    //organisation name
    private static final short ORGANISATION_NAME_ROW_INDEX = 2;
    //organisation legal address
    private static final short ORGANISATION_LEGAL_ADDRESS_ROW_INDEX = 4;
    //organisation factual address
    private static final short ORGANISATION_FACTUAL_ADDRESS_ROW_INDEX = 5;
    //columns
    private static final short HEADER_ROW_INDEX = 14;

    private String organisationName;
    private String organisationLegalAddress;
    private String organisationFactualAddress;
    private static final int HEADER_ROWS_COUNT = 15;

    Columns columns;


    public DocumentHeader(Element documentHeader) {

        //get the name of the organisation
            this.organisationName = documentHeader.getElementsByTagName(en_tag.ORGANISATION_NAME.toString())
                .item(0)
                .getTextContent();

        //get the organisation factual address
        this.organisationFactualAddress = documentHeader.getElementsByTagName(en_tag.ORGANISATION_FACTUAL_ADDRESS.toString())
                .item(0)
                .getTextContent();

        //get the organisation legal address
        this.organisationLegalAddress = documentHeader.getElementsByTagName(en_tag.ORGANISATION_LEGAL_ADDRESS.toString())
                .item(0)
                .getTextContent();

    }

    public DocumentHeader(Sheet sheet, int headerRowIndex) throws IOException {
        //get all existing columns
        Row headerRow = sheet.getRow(headerRowIndex);

        columns = new Columns(headerRow);


    }

    short writeToXLSX(Sheet mainSheet, short currentRow, cl_styles stylesCache) {
        //output the rows of the header
        for (int i = 0; i < HEADER_ROWS_COUNT; i++) {


            if (i == ORGANISATION_NAME_ROW_INDEX) {
                //output organisation name
                Row organisationNameRow = mainSheet.createRow(i);
                Cell organisationNameCell = organisationNameRow.createCell(ColumnIndex.ORGANISATION_NAME);
                organisationNameCell.setCellValue(this.organisationName);
                //merge cells from photo till sum
                mainSheet.addMergedRegion(new CellRangeAddress(i, i, ColumnIndex.HEADER_PHOTO, ColumnIndex.HEADER_SUM));


            } else if (i == ORGANISATION_LEGAL_ADDRESS_ROW_INDEX) {
                //output organisation legal address
                Row organisationLegalAddressRow = mainSheet.createRow(i);
                Cell organisationLegalAddresaCell = organisationLegalAddressRow.createCell(ColumnIndex.ORGANISATION_LEGAL_ADDRESS);
                organisationLegalAddresaCell.setCellValue(this.organisationLegalAddress);
                //merge cells from photo till sum
                mainSheet.addMergedRegion(new CellRangeAddress(i, i, ColumnIndex.HEADER_PHOTO, ColumnIndex.HEADER_SUM));

            } else if (i == ORGANISATION_FACTUAL_ADDRESS_ROW_INDEX) {
                //output organisation factual address
                Row organisationFactualAddressRow = mainSheet.createRow(i);
                Cell organisationFactualAddressCell = organisationFactualAddressRow.createCell(ColumnIndex.ORGANISATION_FACTUAL_ADDRESS);
                organisationFactualAddressCell.setCellValue(this.organisationFactualAddress);
                //merge cells from photo till sum
                mainSheet.addMergedRegion(new CellRangeAddress(i, i, ColumnIndex.HEADER_PHOTO, ColumnIndex.HEADER_SUM));

            } else if (i == HEADER_ROW_INDEX) {
                //get the header row
                Row headerRow = mainSheet.createRow(i);
                //get the header style
                CellStyle headerCellStyle = stylesCache.getHeaderCellStyle();

                //show photo cell
                Cell headerPhotoCell = headerRow.createCell(ColumnIndex.HEADER_PHOTO);
                headerPhotoCell.setCellValue(ColumnHeaderTexts.PHOTO);
                headerPhotoCell.setCellStyle(headerCellStyle);

                //show vendor code
                Cell headerVendorCodeCell = headerRow.createCell(ColumnIndex.HEADER_VENDOR_CODE);
                headerVendorCodeCell.setCellValue(ColumnHeaderTexts.VENDOR_CODE);
                headerVendorCodeCell.setCellStyle(headerCellStyle);

                //show material name header
                Cell headerDescriptionCell = headerRow.createCell(ColumnIndex.HEADER_MATERIAL_NAME);
                headerDescriptionCell.setCellValue(ColumnHeaderTexts.MATERIAL_NAME);
                headerDescriptionCell.setCellStyle(headerCellStyle);

                //show barcode header
                Cell barcodeCell = headerRow.createCell(ColumnIndex.HEADER_BARCODE);
                barcodeCell.setCellValue(ColumnHeaderTexts.BARCODE);
                barcodeCell.setCellStyle(headerCellStyle);

                //show size header
                Cell headerSizeCell = headerRow.createCell(ColumnIndex.HEADER_ATTRIBUTE_NAME);
                headerSizeCell.setCellValue(ColumnHeaderTexts.ATTRIBUTE_NAME);
                headerSizeCell.setCellStyle(headerCellStyle);

                //show santimeters header
                Cell headerSantimetersCell = headerRow.createCell(ColumnIndex.HEADER_SANTIMETERS);
                headerSantimetersCell.setCellValue(ColumnHeaderTexts.SANTIMETERS);
                headerSantimetersCell.setCellStyle(headerCellStyle);

                //show price
                Cell headerPriceCell = headerRow.createCell(ColumnIndex.HEADER_PRICE);
                headerPriceCell.setCellValue(ColumnHeaderTexts.PRICE);
                headerPriceCell.setCellStyle(headerCellStyle);

                //show discount
                Cell headerDiscountCell = headerRow.createCell(ColumnIndex.HEADER_DISCOUNT);
                headerDiscountCell.setCellValue(ColumnHeaderTexts.DISCOUNT);
                headerDiscountCell.setCellStyle(headerCellStyle);

                //show price with discount
                Cell headerPriceWithDiscount = headerRow.createCell(ColumnIndex.HEADER_PRICE_WITH_DISCOUNT);
                headerPriceWithDiscount.setCellValue(ColumnHeaderTexts.PRICE_WITH_DISCOUNT);
                headerPriceWithDiscount.setCellStyle(headerCellStyle);

                //show quantity
                Cell headerQuantityCell = headerRow.createCell(ColumnIndex.HEADER_QUANTITY);
                headerQuantityCell.setCellValue(ColumnHeaderTexts.QUANTITY);
                headerQuantityCell.setCellStyle(headerCellStyle);

                //show you order
                Cell headerYourOrderCell = headerRow.createCell(ColumnIndex.HEADER_SOLD_QUANTITY);
                headerYourOrderCell.setCellValue(ColumnHeaderTexts.SOLD_QUANTITY);
                headerYourOrderCell.setCellStyle(headerCellStyle);

                //show  sum cell
                Cell headerSumCell = headerRow.createCell(ColumnIndex.HEADER_SUM);
                headerSumCell.setCellValue(ColumnHeaderTexts.SUM);
                headerSumCell.setCellStyle(headerCellStyle);
            }
            currentRow++;

        }

        return currentRow;
    }

    public void writeToXML(Element rootElement, Document doc) {
    }

    public Columns getColumns() {
        return columns;
    }
}