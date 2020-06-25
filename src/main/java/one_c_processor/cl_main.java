package one_c_processor;


import one_c_processor.XLSX2XMLConverter.cl_xlsx_to_xml_converter;
import one_c_processor.XML2XLSXConverter.PictureParser;
import one_c_processor.XML2XLSXConverter.XML2XLSXConverter;
import one_c_processor.upd_converter.cl_xml_to_upp_converter;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;


import java.util.HashMap;
import java.util.List;

public final class cl_main {
    private HashMap<String, if_command> actions;

    private cl_main() {
        actions = new HashMap<String, if_command>();
        actions.put(cl_actions.GET_XLSX_FROM_XML, new cl_xlsx_to_xml_converter());
        actions.put(cl_actions.GET_XML_FROM_XLSX, new XML2XLSXConverter());
        actions.put(cl_actions.GENERATE_PICTURES, new PictureParser());
        actions.put(cl_actions.PROCESS_UPD_WILDBERRIES, new cl_xml_to_upp_converter());
    }

    private void run(String[] args) {
        try {

            Path path = Paths.get(args[0]);
            List<String> parametersList = Files.readAllLines(path, StandardCharsets.UTF_8);
            System.out.println(parametersList.size());
            for (int i = 0; i < parametersList.size(); i++) {
                System.out.println(parametersList.get(i));
            }

            //depending on the action instantiate different files
            String action = parametersList.get(1);//the first argument is always action
            actions.get(action).execute(parametersList);

            if (cl_actions.GET_XLSX_FROM_XML.toString().equals(action)) {

                //we are requested to translate the XMl file into XLSX document. instantiate the translator
                //define source folder

                String sourceXML = parametersList.get(2);
                String targetExcel = parametersList.get(3);

                XML2XLSXConverter xml2XLSXConverter = XML2XLSXConverter.createConverter(sourceXML, targetExcel);
                xml2XLSXConverter.convert();
            } else if (cl_actions.GET_XML_FROM_XLSX.toString().equals(action)) {



            } else if (cl_actions.GENERATE_PICTURES.toString().equals(action)) {

            }


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        new cl_main().run(args);

    }
}