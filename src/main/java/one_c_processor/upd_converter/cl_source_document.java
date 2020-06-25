package one_c_processor.upd_converter;

import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.util.HashMap;

public class cl_source_document {

    private HashMap<String, String> material_code_dictionary;

    public cl_source_document(String iv_path_to_source_document) throws Exception {
        File xmlFile = new File(iv_path_to_source_document);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = null;

        dBuilder = factory.newDocumentBuilder();
        var doc = dBuilder.parse(xmlFile);
        NodeList materials = doc.getElementsByTagName("row");
        material_code_dictionary = new HashMap<String, String>();
        int materials_count = materials.getLength();
        for (int i = 0; i < materials_count; i++) {
            add_element_to_hash_map((Element) materials.item(i));
        }


    }


    public String get_barcode_for_code_characteristic(String iv_line_number) {
        return material_code_dictionary.get(iv_line_number);
    }

    private void add_element_to_hash_map(Element material) {
        var nodelist = material.getElementsByTagName("Value");
        var node_list_length = nodelist.getLength();
        var lv_record_number = nodelist.item(0).getTextContent();
        var lv_material_barcode = nodelist.item(1).getTextContent();

        material_code_dictionary.put(lv_record_number,lv_material_barcode);
    }


}

