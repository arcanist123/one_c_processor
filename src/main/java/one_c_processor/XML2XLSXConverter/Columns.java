package one_c_processor.XML2XLSXConverter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.*;
import java.util.function.Consumer;

public class Columns implements Iterable<cl_column> {

    public static final String materialDescription = "MaterialDescription";
    public static final String materialVendorCode = "MaterialVendorCode";
    public static final String attributeDescription = "AttributeDescription";
    public static final String attributeSantimeters = "AttributeSantimeters";
    public static final String attributePrice = "AttributePrice";
    public static final String attributeDiscount = "AttributeDiscount";
    public static final String attributePriceWithDiscount = "AttributePriceWithDiscount";
    public static final String attributeQuantity = "AttributeQuantity";
    public static final String attributeSoldQuantity = "AttributeSoldQuantity";
    public static final String materialCountryOfOrigin = "MaterialCountryOfOrigin";
    public static final String materialComposition = "MaterialComposition";
    public static final String materialGroup = "MaterialGroup";
    public static final String attributeSum = "AttributeSum";
    public static final String originalRow = "OriginalRow";
    public static final String materialPicture = "MaterialPicture";

    private ArrayList<cl_column> columns;
    private static final String HEADER_MATERIAL_GROUP_GUID = "GUID группы";
    private static final String HEADER_MATERIAL_GUID = "GUID товара";
    private static final String HEADER_MATERIAL_ATTRIBUTE_GUID = "GUID характеристики";
    private static final String HEADER_PHOTO = "Фото";
    private static final String HEADER_MATERIAL_VENDOR_CODE = "Артикул";
    private static final String HEADER_MATERIAL_DESCRIPTION = "Описание";
    private static final String HEADER_ATTRIBUTE_DESCRIPTION = "Размер";
    private static final String HEADER_ATTRIBUTE_SANTIMETERS = "Сантиметры";
    private static final String HEADER_ATTRIBUTE_PRICE = "Цена";
    private static final String HEADER_DISCOUNT = "Скидка";
    private static final String HEADER_PRICE_WITH_DISCOUNT = "Цена со скидкой";
    private static final String HEADER_QUANTITY = "Количество";
    private static final String HEADER_SOLD_QUANTITY = "Ваш Заказ";
    private static final String HEADER_SUM = "Итого";
    private static final String HEADER_COUNTRY = "Страна";
    private static final String HEADER_COMPOSITION = "Состав";
    private static final String HEADER_MATERIAL_GROUP = "Группа";


    @Override
    public Iterator<cl_column> iterator() {
        return columns.iterator();
    }

    @Override
    public void forEach(Consumer<? super cl_column> action) {

    }

    @Override
    public Spliterator<cl_column> spliterator() {
        return null;
    }

    public cl_column getColumn(String columnTag) {
        for (cl_column column : columns) {
            if (column.getTag() == columnTag) {
                return column;
            }
        }
        return null;
    }

    public Columns(Row headerRow) {
        columns = new ArrayList<cl_column>();
        for (Cell cell : headerRow) {
            //check all string cells
            if (cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getStringCellValue().toUpperCase().equals(HEADER_MATERIAL_VENDOR_CODE.toUpperCase())) {
                    this.columns.add(new cl_column(materialVendorCode, HEADER_MATERIAL_VENDOR_CODE, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_MATERIAL_DESCRIPTION.toUpperCase())) {
                    this.columns.add(new cl_column(materialDescription, HEADER_MATERIAL_DESCRIPTION, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_ATTRIBUTE_DESCRIPTION.toUpperCase())) {
                    this.columns.add(new cl_column(attributeDescription, HEADER_ATTRIBUTE_DESCRIPTION, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_ATTRIBUTE_SANTIMETERS.toUpperCase())) {
                    this.columns.add(new cl_column(attributeSantimeters, HEADER_ATTRIBUTE_SANTIMETERS, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_ATTRIBUTE_PRICE.toUpperCase())) {
                    this.columns.add(new cl_column(attributePrice, HEADER_ATTRIBUTE_PRICE, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_DISCOUNT.toUpperCase())) {
                    this.columns.add(new cl_column(attributeDiscount, HEADER_DISCOUNT, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_PRICE_WITH_DISCOUNT.toUpperCase())) {
                    this.columns.add(new cl_column(attributePriceWithDiscount, HEADER_PRICE_WITH_DISCOUNT, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_QUANTITY.toUpperCase())) {
                    this.columns.add(new cl_column(attributeQuantity, HEADER_QUANTITY, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_SOLD_QUANTITY.toUpperCase())) {
                    this.columns.add(new cl_column(attributeSoldQuantity, HEADER_SOLD_QUANTITY, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_COUNTRY.toUpperCase())) {
                    this.columns.add(new cl_column(materialCountryOfOrigin, HEADER_COUNTRY, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_COMPOSITION.toUpperCase())) {
                    this.columns.add(new cl_column(materialComposition, HEADER_COMPOSITION, cell.getColumnIndex()));

                } else if (cell.getStringCellValue().toUpperCase().equals(HEADER_MATERIAL_GROUP.toUpperCase())) {
                    this.columns.add(new cl_column(materialGroup, HEADER_MATERIAL_GROUP, cell.getColumnIndex()));
                }
            }
        }
    }
}