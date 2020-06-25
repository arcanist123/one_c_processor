package one_c_processor.XML2XLSXConverter;

import one_c_processor.en_tag;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.IOException;
import java.util.Vector;

public class MaterialGroup {
    private static final short GROUP_GUID_COLUMN = 0;
    private static final short GROUP_NAME_COLUMN = 3;


    private String groupName;
    private String groupGUID;

    private Vector<Material> materials;

    private MaterialGroup() {
        this.materials = new Vector<Material>(0);//instantiate an empty vector of materials
    }

    //public section
    public static MaterialGroup createMaterialGroup(Element materialGroupElement) throws IOException {
        //create the instance of material group
        MaterialGroup materialGroup = new MaterialGroup();

        //get the name of the material group
        materialGroup.groupName = materialGroupElement.getElementsByTagName(en_tag.GROUP_NAME.toString()).item(0).getTextContent();

        //get the material group GUID
        materialGroup.groupGUID = materialGroupElement.getElementsByTagName(en_tag.GROUP_GUID.toString()).item(0).getTextContent();

        //parse the individual materials into  vector of materials
        NodeList materials = materialGroupElement.getElementsByTagName(en_tag.MATERIAL.toString());

        for (int i = 0; i < materials.getLength(); i++) {
            //get current material as element
            Element materialElement = (Element) materials.item(i);

            //convert the current material into the Material instance and add it into the vector of materials
            materialGroup.materials.add(Material.createMaterial(materialElement));
        }

        return materialGroup;
    }


    public short writeToXLSX(Sheet mainSheet, short currentRow, cl_styles stylesCache) {
        // write the header row of the material group
        // Create a row for the material group header
        Row row = mainSheet.createRow(currentRow);
        short startRowExcel = currentRow; //preserve the starting row of the group
        startRowExcel++;//add 1 to start row to show it in the excel notation

        //write the guid of the row
        Cell cellGroupGuid = row.createCell(GROUP_GUID_COLUMN);
        cellGroupGuid.setCellValue(this.groupGUID);

        //write the name of the group
        Cell cellGroupName = row.createCell(GROUP_NAME_COLUMN);
        cellGroupName.setCellValue(this.groupName);

        //get the cell style for the column group
        CellStyle style = stylesCache.getGroupCellStyle();
        cellGroupName.setCellStyle(style);

        //increase the index to show that we are done with this row
        currentRow++;

        //output all the materials
        for (int i = 0; i < this.materials.size(); i++) {
            currentRow = this.materials.elementAt(i).writeToXLSX(mainSheet, currentRow, stylesCache);

        }

        //output group guid for all the
        for (int i = cellGroupGuid.getRowIndex() + 1; i < currentRow; i++) {
            Row materialRow = mainSheet.getRow(i);
            Cell materialGroupGuidCell = materialRow.createCell(GROUP_GUID_COLUMN);
            materialGroupGuidCell.setCellValue(this.groupGUID);
        }

        //after we have written all the materials, group them into one group
        mainSheet.groupRow(startRowExcel, currentRow - 1);

        //increase the index to make the empty row after we have written the group
        currentRow++;

        //return the current index
        return currentRow;
    }


}
