package one_c_processor.upd_converter;


import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;

public class cl_source_upd {
    private Document doc;

    public cl_source_upd(String iv_upd_source_path) {
        File xmlFile = new File(iv_upd_source_path);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = null;
        try {
            dBuilder = factory.newDocumentBuilder();
            this.doc = dBuilder.parse(xmlFile);

        } catch (ParserConfigurationException | SAXException | IOException e) {
            e.printStackTrace();
        }
    }


    public Document get_adjusted_upd(cl_source_document io_source_document) {
        var materials = doc.getElementsByTagName("СведТов");
        var materials_length = materials.getLength();
        for (var i = 0; i < materials_length; i++) {
            var material = (Element) materials.item(i);
            var line_number = material.getAttribute("НомСтр");
            var additional_informations = material.getElementsByTagName("ДопСведТов");
            assert additional_informations.getLength() == 1;
            var additional_information = (Element) additional_informations.item(0);
            var barcode = io_source_document.get_barcode_for_code_characteristic(line_number);
            additional_information.setAttribute("КодТов", barcode);
        }

        return doc;
    }

}
