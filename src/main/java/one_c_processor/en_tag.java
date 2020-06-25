package one_c_processor;


public enum en_tag {
    ORGANISATION("Organisation"),
    MATERIAL("Material"),
    MATERIAL_NAME("MaterialName"),
    MATERIAL_GUID("MaterialGUID"),
    ATTRIBUTE("Attribute"),
    ATTRIBUTE_GUID("AttributeGUID"),
    ATTRIBUTE_PRICE("AttributePrice"),
    ATTRIBUTE_DISCOUNT("AttributeDiscount"),
    ATTRIBUTE_PRICE_WITH_DISCOUNT("AttributePriceWithDiscount"),
    GROUP_GUID("MaterialGroupGUID"),
    GROUP_NAME("MaterialGroupName"),
    MATERIAL_VENDOR_CODE("MaterialVendorCode"),
    PRICE_LIST("PriceList"),
    SANTIMETERS("AttributeSantimeters"),
    PICTURE("MaterialPicture"),
    ORGANISATION_NAME("OrganisationName"),
    ORGANISATION_FACTUAL_ADDRESS("OrganisationFactualAddress"),
    ORGANISATION_LEGAL_ADDRESS("OrganisationLegalAddress"),
    ATTRIBUTE_SOLD_QUANTITY("AttributeSoldQuantity"),
    ATTRIBUTE_QUANTITY("AttributeQuantity"),
    MATERIAL_GROUP("MaterialGroup"),
    MATERIALS_SECTION("Materials"),
    ATTRIBUTE_NAME("AttributeName"),
    BARCODE("Barcode");
    String tagName;

    // Constructor
    en_tag(String mass) {
        this.tagName = mass;

    }


    @Override
    public String toString() {
        return this.tagName;
    }


}