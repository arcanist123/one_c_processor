package one_c_processor.upd_converter;

import one_c_processor.if_command;
import org.w3c.dom.Document;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.util.List;

public class cl_xml_to_upp_converter implements if_command {


    public cl_xml_to_upp_converter() {
    }

    public void convert(String iv_path_to_upd, String iv_path_to_source_document) throws Exception {
        var lo_source_upd = new cl_source_upd(iv_path_to_upd);
        var lo_source_document = new cl_source_document(iv_path_to_source_document);
        var lo_adjusted_upd = lo_source_upd.get_adjusted_upd(lo_source_document);
        this.save_adjusted_upd(lo_adjusted_upd, iv_path_to_upd);


    }

    private void save_adjusted_upd(Document lo_adjusted_upd, String iv_path_to_upd) throws Exception {
        //the XMl document was formulated. Now we can write it to the file
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
        DOMSource source = new DOMSource(lo_adjusted_upd);
        StreamResult result = new StreamResult(new File(iv_path_to_upd));
        transformer.transform(source, result);
    }

    @Override
    public void execute(List<String> it_parameters) throws Exception {
        var lv_path_to_upd = it_parameters.get(2);
        var lv_path_to_source_document = it_parameters.get(3);
        new cl_xml_to_upp_converter().convert(lv_path_to_upd, lv_path_to_source_document);
    }
}
