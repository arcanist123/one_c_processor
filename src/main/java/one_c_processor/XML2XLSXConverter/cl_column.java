package one_c_processor.XML2XLSXConverter;


public class cl_column {
    private String tag;
    private String headerText;
    private int columnIndex;
    public cl_column(String tag, String headerText, int columnIndex){
        this.tag = tag;
        this.columnIndex = columnIndex;
        this.headerText = headerText;
    }
    public String getTag(){
        return tag;
    }

    public int getColumnIndex() {
        return columnIndex;
    }
}
