package one_c_processor.XML2XLSXConverter;


import one_c_processor.if_command;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Hashtable;
import java.util.List;

public class PictureParser implements if_command {
    private Hashtable<Integer, XSSFPictureData> pictureDataMap;

    public PictureParser() {
        pictureDataMap = new Hashtable<>();
    }

    public static PictureParser createParser(Sheet sheet) {

        PictureParser pictureParser = new PictureParser();
        //get the pictures of the list
        Drawing drawing = sheet.getDrawingPatriarch();
        List<XSSFShape> pics = ((XSSFDrawing) drawing).getShapes();

        //process the drawings
        for (XSSFShape pic : pics) {
            XSSFPicture picture = (XSSFPicture) pic;
            XSSFClientAnchor clientAnchor = picture.getClientAnchor();

            pictureParser.pictureDataMap.put(clientAnchor.getRow1(), picture.getPictureData());
        }
        return pictureParser;
    }

    public void convert(String sourceFile, String targetFile) throws IOException {

        //get the file
        Workbook wb = new XSSFWorkbook(sourceFile);
        //get sheet
        Sheet sheet = wb.getSheetAt(0);

        //get the column with Factory code
        int factoryCodeColumnIndex = getFactoryCodeColumnIndex(sheet);

        Drawing drawing = sheet.getDrawingPatriarch();
        List<XSSFShape> pics = ((XSSFDrawing) drawing).getShapes();

        for (XSSFShape pic : pics) {
            //suppress cast conversion exception
            try {
                XSSFPicture inpPic = (XSSFPicture) pic;

                XSSFClientAnchor clientAnchor = inpPic.getClientAnchor();
                System.out.println("col1: " + clientAnchor.getCol1() + ", col2: " + clientAnchor.getCol2() + ", row1: " + clientAnchor.getRow1() + ", row2: " + clientAnchor.getRow2());
                System.out.println("x1: " + clientAnchor.getDx1() + ", x2: " + clientAnchor.getDx2() + ", y1: " + clientAnchor.getDy1() + ", y2: " + clientAnchor.getDy2());
                final XSSFPictureData pictureData = inpPic.getPictureData();

                //get the name of the picture
                //get factory code
                String factoryCode = getFactoryCode(sheet, clientAnchor.getRow1(), factoryCodeColumnIndex);
                //get the file extension
                String fileExtension = pictureData.suggestFileExtension();
                //get the full file name
                String fullFileName = targetFile + "\\" + factoryCode + "." + fileExtension;

                //generate the file
                FileOutputStream out = new FileOutputStream(fullFileName);
                out.write(pictureData.getData());
                out.close();

            } catch (ClassCastException ex) {

            }

        }
    }

    private String getFactoryCode(Sheet sheet, int row1, int factoryColumnIndex) {
        //the factory code is located in the cell with coordingates
        Row factoryRow = sheet.getRow(row1);
        Cell factoryCodeCell = factoryRow.getCell(factoryColumnIndex);

        CellType cellType = factoryCodeCell.getCellTypeEnum();

        if (cellType == CellType.NUMERIC) {
            DecimalFormat decimalFormat = new DecimalFormat("#");
            String numberAsString = decimalFormat.format(factoryCodeCell.getNumericCellValue());
            System.out.println(numberAsString);
            return numberAsString;
        } else {

            return factoryCodeCell.getStringCellValue().trim();
        }
    }

    private int getPictureColumnIndex(Sheet sheet) {
        int pictureColumnIndex = 0;
        //check first 100 cells of the first row
        Row firstRow = sheet.getRow(0);
        for (int i = 0; i < 100; i++) {
            //get the current cell
            Cell currentCell = firstRow.getCell(i);


            //get the value of the cell
            if (currentCell.getStringCellValue() == "Картинка") {
                pictureColumnIndex = i;
            }
        }
        return pictureColumnIndex;

    }

    public int getFactoryCodeColumnIndex(Sheet sheet) {
        int codeColumnIndex = 0;
        //check first 100 cells of the first row
        Row firstRow = sheet.getRow(0);
        for (int i = 0; i < 100; i++) {
            //get the current cell
            Cell currentCell = firstRow.getCell(i);


            //get the value of the cell
            if (currentCell != null) {
                //inform about the type of the current cell
                System.out.println("Cell " + i + " has type " + currentCell.getCellTypeEnum().toString());

                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                    //inform about the value of the current cell
                    System.out.println("Cell " + i + " has value " + currentCell.getStringCellValue().toString());
                    if (currentCell.getStringCellValue().equalsIgnoreCase("Артикул")) {
                        codeColumnIndex = i;
                    }
                }

            }
        }

        return codeColumnIndex;
    }

    public PictureData getPictureData(int i) {
        return pictureDataMap.get(i);
    }


    @Override
    public void execute(List<String> it_parameters) throws Exception {
        //we need to generate the pictures based on the file
        String sourceFile = it_parameters.get(2);
        String targetPath = it_parameters.get(3);

        PictureParser pictureParser = new PictureParser();
        pictureParser.convert(sourceFile, targetPath);
    }
}
