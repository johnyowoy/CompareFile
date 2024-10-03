package Config;

import java.util.ArrayList;
import java.util.HashSet;

public class File {
    // 指定來源Excel檔案
    private String ExcelFileName;
    // 指定來源資料Rows
    private int[] Rows;
    // 收集來源欄位名稱
    private final HashSet<ArrayList<String>> SourceRowName;
    // 收集來源資料
    private final HashSet<ArrayList<String>> SourceData;

    // Get ExcelFileName
    public String getExcelFileName() {
        return ExcelFileName;
    }
    // Set ExcelFileName
    public void setExcelFileName(String addExcelFileName) {
        this.ExcelFileName = addExcelFileName;
    }
    // Get Rows
    public int[] getRows() {
        return Rows;
    }
    // Set Rows
    public void setRows(int[] addRows) {
        this.Rows = addRows;
    }

    public File() {
        // 收集Excel資料，使用HashSet去重複行數
        SourceData = new HashSet<>();
        // 收集Excel標題資料，使用HashSet去重複行數
        SourceRowName = new HashSet<>();
    }
    // Get SourceRowName
    public HashSet<ArrayList<String>> getSourceRowName() {
        return SourceRowName;
    }
    // Get SourceData
    public HashSet<ArrayList<String>> getSourceData() {
        return SourceData;
    }

}
