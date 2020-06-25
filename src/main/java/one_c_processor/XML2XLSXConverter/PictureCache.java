package one_c_processor.XML2XLSXConverter;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

public class PictureCache {
    private static PictureCache instance = null;
    Map<Integer, Integer> map = new HashMap<Integer, Integer>();

    private PictureCache() {
        // Exists only to defeat instantiation.
    }

    public static PictureCache getInstance() {
        if (instance == null) {
            instance = new PictureCache();
        }
        return instance;
    }


    public int addPicture(Workbook workbook, byte[] picture, int pictureType) {
        //get the hash from array
        final int hashCode = Arrays.hashCode(picture);

        int pictureId;//will serve as a temp var for picture id

        //check if picture was already added into the workbook
        if (map.containsKey(hashCode) == true) {
            //such a picture was already added into the workbook. get the index of the picture
            pictureId = map.get(hashCode);
        } else {
            //the picture was not already added into the workbook. add it and store into the cache
            pictureId = workbook.addPicture(picture, pictureType);
            map.put(hashCode, pictureId);
        }
        //return the id of the picture
        return pictureId;

    }
}
