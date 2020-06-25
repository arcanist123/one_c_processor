package one_c_processor.XML2XLSXConverter;


import one_c_processor.if_command;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class XML2XLSXConverter implements if_command {
    private static XML2XLSXConverter instance = null;
    private String sourceXML;
    private String targetExcel;
    private String constSourceFileName = "source.xml";


    public XML2XLSXConverter() {
        // Exists only to defeat instantiation.
    }


    public static XML2XLSXConverter getInstance() {
        if (instance == null) {
            instance = new XML2XLSXConverter();
        }
        return instance;
    }

    public static XML2XLSXConverter createConverter(String sourceXML, String targetExcel) {
        XML2XLSXConverter instance = XML2XLSXConverter.getInstance();
        instance.sourceXML = sourceXML;
        instance.targetExcel = targetExcel;
        return instance;
    }

    private static PriceList getConvertedDocument(String sourceXML) throws ParserConfigurationException, SAXException, IOException {
        //get the name of the source file. By convention, it is named source.xml and is located in the source folder
        File xmlFile = new File(sourceXML);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = factory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);

        //create the excel document Java representation
        PriceList priceList = new PriceList(doc.getDocumentElement());
        return priceList;
    }

    public String getSourceXML() {
        return sourceXML;
    }

    public void convert() throws ParserConfigurationException, IOException, SAXException {
        //price list
        PriceList priceList = getConvertedDocument(this.sourceXML);

        //we have the instance of all materials.
        //create the xlsx document
        Workbook wb = new XSSFWorkbook();
        Sheet mainSheet = wb.createSheet("ПРАЙС");
        //ensure that the grouping in this sheet is on top
        mainSheet.setRowSumsBelow(false);

        //write the PriceList to xlsx
        priceList.writeToXLSX(mainSheet);

        //write the xlsx Price list into the file
        FileOutputStream fileOut = new FileOutputStream(this.targetExcel);
        wb.write(fileOut);
        fileOut.close();

    }

    @Override
    public void execute(List<String> it_parameters) throws Exception {
        String sourceXML = it_parameters.get(2);
        String targetExcel = it_parameters.get(3);

        XML2XLSXConverter xml2XLSXConverter = XML2XLSXConverter.createConverter(sourceXML, targetExcel);
        xml2XLSXConverter.convert();

    }
}