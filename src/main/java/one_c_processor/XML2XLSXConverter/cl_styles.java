package one_c_processor.XML2XLSXConverter;


import org.apache.poi.ss.usermodel.*;

public class cl_styles {
    private static cl_styles instance = null;
    private CellStyle groupCellStyle = null;

    private Workbook wb;
    private CellStyle materialNameCellStyle = null;
    private CellStyle materialVendorCodeCellStyle = null;
    private CellStyle attributeNameCellStyle = null;
    private CellStyle attributePriceCellStyle = null;
    private CellStyle attributeDiscountCellStyle;
    private CellStyle attributePriceWithDiscountCellStyle = null;
    private CellStyle headerCellStyle = null;
    private CellStyle quantityCellStyle = null;
    private CellStyle sumCellStyle = null;
    private short rublesFormat = 0;
    private CellStyle attributeOrderCellStyle;

    protected cl_styles() {
        // Exists only to defeat instantiation.
    }


    public static cl_styles createInstance(Workbook wb) {

        if (instance == null) {
            instance = new cl_styles();
        }
        instance.wb = wb;
        return instance;

    }


    public CellStyle getGroupCellStyle() {


        //instantiate the style in case it was not already instantiated
        if (groupCellStyle == null) {
            //we do not have the style already. create it

            //create the font
            Font font = wb.createFont();
            font.setFontHeightInPoints((short) 12);
            font.setFontName("CONSOLAS");
            font.setBold(true);
            //font.setColor(IndexedColors.BRIGHT_GREEN.index);

            //create the style
            this.groupCellStyle = wb.createCellStyle();
            this.groupCellStyle.setFont(font);
        } else {
            //the style was already created. do nothing
        }

        //return the created style
        return this.groupCellStyle;
    }

    private CellStyle getCommonCellStyle(Workbook wb) {


        //create the font
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("CONSOLAS");


        //create the style and set the font
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);

        //create the borders
        BorderStyle borderStyle = BorderStyle.MEDIUM;
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderLeft(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        cellStyle.setBorderTop(borderStyle);

        //set the alignment
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setWrapText(true);

        return cellStyle;
    }


    public CellStyle getMaterialVendorCodeCellStyle() {
        //instantiate the attribute in case it is not already instantiated
        if (this.materialVendorCodeCellStyle == null) {
            this.materialVendorCodeCellStyle = this.getCommonCellStyle(wb);
        } else {

        }
        return this.materialVendorCodeCellStyle;
    }

    public CellStyle getAttributeDiscountCellStyle() {
        //instantiate the attribute in case it is not already instantiated
        if (this.attributeDiscountCellStyle == null) {
            this.attributeDiscountCellStyle = this.getCommonCellStyle(wb);
        } else {

        }
        return this.attributeDiscountCellStyle;
    }

    public CellStyle getAttributeNameCellStyle() {
        //instantiate the attribute in case it is not already instantiated
        if (this.attributeNameCellStyle == null) {
            this.attributeNameCellStyle = this.getCommonCellStyle(wb);
        } else {

        }
        return this.attributeNameCellStyle;
    }

    public CellStyle getMaterialNameCellStyle() {
        //instantiate the attribute in case it is not already instantiated
        if (this.materialNameCellStyle == null) {
            //the material name cell style is nothing special - get it as a common style
            this.materialNameCellStyle = this.getCommonCellStyle(this.wb);
        } else {

        }
        return this.materialNameCellStyle;
    }

    public CellStyle getAttributePriceCellStyle() {

        if (this.attributePriceCellStyle == null) {
            //the material name cell style is nothing special - get it as a common style
            this.attributePriceCellStyle = this.getCommonCellStyle(this.wb);
            //set the custom format for the cell
            this.attributePriceCellStyle.setDataFormat(this.getRublesFormat());
        } else {
            //the price cell style was already initialised
        }
        return this.attributePriceCellStyle;
    }

    public CellStyle getAttributePriceWithDiscountCellStyle() {

        if (this.attributePriceWithDiscountCellStyle == null) {
            //the material name cell style is nothing special - get it as a common style
            this.attributePriceWithDiscountCellStyle = this.getCommonCellStyle(this.wb);
            //set the custom format for the cell
            this.attributePriceWithDiscountCellStyle.setDataFormat(this.getRublesFormat());
            this.attributePriceWithDiscountCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            this.attributePriceWithDiscountCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        } else {

        }
        return this.attributePriceWithDiscountCellStyle;
    }

    private short getRublesFormat() {
        if (this.rublesFormat == 0) {
            //the format for rubles was not created yet. generate it
            DataFormat format = this.wb.createDataFormat();
            this.rublesFormat = format.getFormat("#,##0\"р.\"");

        } else {
            //no need to do anything - we have already created the format
        }

        return this.rublesFormat;

    }

    public CellStyle getHeaderCellStyle() {
        if (this.headerCellStyle == null) {
            //the material name cell style is nothing special - get it as a common style
            this.headerCellStyle = this.getCommonCellStyle(this.wb);
        } else {

        }
        return this.headerCellStyle;

    }

    public CellStyle getQuantityStyle() {
        if (this.quantityCellStyle == null) {
            //the material name cell style is nothing special - get it as a common style
            this.quantityCellStyle = this.getCommonCellStyle(this.wb);
        } else {

        }
        return this.quantityCellStyle;

    }

    public CellStyle getSumCellStyle() {
        if (this.sumCellStyle == null) {
            //the material name cell style is nothing special - get it as a common style
            this.sumCellStyle = this.getCommonCellStyle(this.wb);
            //set the custom format for the cell
            this.sumCellStyle.setDataFormat(this.getRublesFormat());
        } else {

        }
        return this.sumCellStyle;
    }

    public CellStyle getAttributeOrderCellStyle() {
        if (this.attributeOrderCellStyle == null) {
            //the order is nothing special - get the common style
            this.attributeOrderCellStyle = this.getCommonCellStyle(this.wb);
        } else {
            //we have already created this style. do nothing
        }

        return this.attributeOrderCellStyle;

    }
}