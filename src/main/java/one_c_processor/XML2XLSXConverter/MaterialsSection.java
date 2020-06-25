package one_c_processor.XML2XLSXConverter;

import one_c_processor.en_tag;
import org.apache.poi.ss.usermodel.*;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.IOException;
import java.util.Base64;
import java.util.Vector;

public class MaterialsSection {

    private Vector<MaterialGroup> materialGroupsVector;
    private String sumFormula = "";

    private Sheet sheet;
    private int materialsStartRowIndex;
    private Columns columns;

    private MaterialsSection() {
        this.materialGroupsVector = new Vector<MaterialGroup>(0);
    }

    public MaterialsSection(Sheet sheet, int columnHeaderRowIndex, Columns columns) {
        this.sheet = sheet;
        this.columns = columns;
        this.materialsStartRowIndex = columnHeaderRowIndex + 1;
    }

    public void writeToXML(Element rootElement, Document doc) {

        //output materials and their info
        writeMaterialsToXML(rootElement, doc);
    }

    private void writeMaterialsToXML(Element rootElement, Document doc) {
        Element materials = doc.createElement(en_tag.MATERIALS_SECTION.toString());
        rootElement.appendChild(materials);

        //get the column for vendor code
        int vendorCodeColumnIndex = columns.getColumn(Columns.materialVendorCode).getColumnIndex();
        //get the attribute description
        int attributeDescriptionColumnIndex = columns.getColumn(Columns.attributeDescription).getColumnIndex();

        PictureParser pictureParser = PictureParser.createParser(sheet);
        //process the lines of the document
        for (Row row : sheet) {
            if (row.getRowNum() > this.materialsStartRowIndex) {
                //get vendor code
                Cell vendorCodeCell = row.getCell(vendorCodeColumnIndex);
                //get attribute name
                Cell attributeNameCell = row.getCell(attributeDescriptionColumnIndex);
                if (vendorCodeCell == null || attributeNameCell == null) {
                    //no need to do anything - there is no cell either vendor code or attribute description. Skip this row
                } else {
                    vendorCodeCell.setCellType(CellType.STRING);
                    String vendorCodeCellValue = vendorCodeCell.getStringCellValue();

                    if (vendorCodeCellValue.contains("1023")){
                        System.out.println(vendorCodeCellValue.toString());
                    }

                    attributeNameCell.setCellType(CellType.STRING);
                    String attributeNameCellValue = attributeNameCell.getStringCellValue();


                    if (vendorCodeCellValue == "" || attributeNameCellValue == "") {
                        //this is not a valid row - we don't need to process it
                    } else {
                        Element material = doc.createElement(en_tag.MATERIAL.toString());
                        materials.appendChild(material);

                        //this row has a vendor code and the attribute name. We can write such row into the xml file
                        for (cl_column column : this.columns) {
                            //get cell value
                            Cell cell = row.getCell(column.getColumnIndex());
                            String cellValue = "";
                            //ensure that we are only processing the cell in case the cell is there
                            if (cell != null) {
                                if (cell.getCellTypeEnum() == CellType.STRING) {
                                    cellValue = cell.getStringCellValue();

                                } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cellValue = Double.toString(cell.getNumericCellValue());
                                }
                            }

                            Element element = doc.createElement(column.getTag());
                            element.appendChild(doc.createTextNode(cellValue));
                            material.appendChild(element);
                        }
                        //add the original row number
                        Element element = doc.createElement(Columns.originalRow);
                        String rowNumberString = Double.toString(row.getRowNum() + 1);
                        element.appendChild(doc.createTextNode(rowNumberString));
                        material.appendChild(element);

                        //add the picture
                        element = doc.createElement(Columns.materialPicture);
                        PictureData pictureData = pictureParser.getPictureData(row.getRowNum() );
                        //add the picture in case we have a picture for the current row
                        String pictureDataString = "";
                        if (pictureData != null) {
                            pictureDataString = Base64.getEncoder().encodeToString(pictureData.getData());
                            element.appendChild(doc.createTextNode(pictureDataString));
                        }
                        element.appendChild(doc.createTextNode(pictureDataString));
                        material.appendChild(element);
                    }
                }
            }
        }
    }

    public static MaterialsSection createMaterials(Element materialsElement) throws IOException {
        //create an instance of materialsSection. It actually does not have any attributes and is just a holder for the vector of groups
        MaterialsSection materialsSection = new MaterialsSection();


        //get all the material groups elements
        NodeList materialGroupsList = materialsElement.getElementsByTagName(en_tag.MATERIAL_GROUP.toString());
        for (int i = 0; i < materialGroupsList.getLength(); i++) {
            //get the current material group element
            Element materialGroupElement = (Element) materialGroupsList.item(i);
            //add the current material group to the vector of material groups
            materialsSection.materialGroupsVector.add(MaterialGroup.createMaterialGroup(materialGroupElement));
        }

        //return the result of processing - the materialsSection instance
        return materialsSection;

    }

    public short writeToXLSX(Sheet mainSheet, short currentRow, cl_styles stylesCache) {
        //we need to output all the groups we have
        this.initSumFormula();
        for (int i = 0; i < this.materialGroupsVector.size(); i++) {
            //output the current group
            short currentMaterialGroupStart = currentRow;
            currentRow = this.materialGroupsVector.elementAt(i).writeToXLSX(mainSheet, currentRow, stylesCache);
            this.updateSumFormula(currentMaterialGroupStart, currentRow);
        }
        this.closeSumFormula();


        //we have finished the output of all the groups. Now output the sum of all of them
        this.writeSumFormula(mainSheet, currentRow, stylesCache);

        //return the current index in the xlsx
        return currentRow;
    }

    private void closeSumFormula() {
        sumFormula = sumFormula + ")";
    }

    private void initSumFormula() {
        sumFormula = "SUM(";
    }

    private void writeSumFormula(Sheet mainSheet, short currentRow, cl_styles stylesCache) {
        //output the header to the sum
        Row row = mainSheet.createRow(currentRow - 1);
        Cell headerSumCell = row.createCell(ColumnIndex.HEADER_SOLD_QUANTITY);
        headerSumCell.setCellValue("Сумма:");
        CellStyle headerCellStyle = stylesCache.getHeaderCellStyle();
        headerSumCell.setCellStyle(headerCellStyle);

        //output the actual formula
        Cell valueSumCell = row.createCell(ColumnIndex.HEADER_SUM);
        valueSumCell.setCellFormula(this.sumFormula);
        CellStyle attributeSumCellStyle = stylesCache.getSumCellStyle();
        valueSumCell.setCellStyle(attributeSumCellStyle);
    }

    private void updateSumFormula(short currentMaterialGroupStart, short currentMaterialGroupEnd) {
        //update the sum string with the current formula
        short start = (short) (currentMaterialGroupStart + 2);
        short end = (short) (currentMaterialGroupEnd - 1);
        if (sumFormula == "SUM(") {
            sumFormula = sumFormula + "N" + start + ":" + "N" + end;
        } else {
            sumFormula = sumFormula + "," + "N" + start + ":" + "N" + end;
        }
    }
}
