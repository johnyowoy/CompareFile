package Config;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

public class Excel extends File {
    // 收集來源資料並且去重複行數，Action
    public void CollectionStore() {
        try (InputStream file = new FileInputStream("resources\\SourceData\\" + getExcelFileName() + ".xlsx")) {

            Workbook wb = WorkbookFactory.create(file);
            Sheet sheet = wb.getSheetAt(0);
            // 取得Excel總行数，共有多少筆資料
            int rowCount = sheet.getLastRowNum()+1;

            // 取得欄位名稱
            Row RowsName = sheet.getRow(0);
            // 建立空的陣列，將欄位名稱塞進去
            ArrayList<String> rowNameData = new ArrayList<>();
            for (int cellIndex : getRows()) {
                Cell cell = RowsName.getCell(cellIndex);
                String cellString = (cell != null) ? getValueFromCell(cell) : "null";
                rowNameData.add(cellString);
            }
            getSourceRowName().add(rowNameData);

            // 循環每一行並且收集指定的資料
            // 從工作表第二行開始(因為第一行是標題)，所以i=1
            for (int i = 1; i < rowCount; i++) {
                Row row = sheet.getRow(i);
                // 建立空的陣列，將資料塞進去
                ArrayList<String> rowData = new ArrayList<>();
                //rowData.add(getExcelFileName());
                for (int cellIndex : getRows()) {
                    Cell cell = row.getCell(cellIndex);
                    String cellString = (cell != null) ? getValueFromCell(cell) : "null";
                    // 移除特殊字元
                    cellString = cellString.replaceAll("test", "");
                    rowData.add(cellString);
                }
                // 每行有相同的資料，進行去重複
                getSourceData().add(rowData);
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    // 顯示SourceData內容
    public void showSourceData() {
        // 檢查 SourceData是否有資料
        if (getSourceData().isEmpty()) {
            System.out.println("No data collected.");
        } else {
            for (ArrayList<String> rowData : getSourceRowName()) {
                System.out.println(rowData);
            }
            for (ArrayList<String> rowData : getSourceData()) {
                System.out.println(rowData);
            }
        }
    }

    private static String getValueFromCell(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return String.valueOf(cell.getCellFormula());
            default:
                return "";
        }
    }


}