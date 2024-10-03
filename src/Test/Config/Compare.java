package Config;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;

public class Compare {
    private Excel[] files;
    private Workbook workbook;

    public void setFile(Excel[] files) {
        this.files = files;
        // 建立新的Excel Workbook
        this.workbook = new XSSFWorkbook();
    }

    // 比較兩個Excel檔案
    public void compareFiles() {
        if (files == null || files.length < 2) {
            System.out.println("需要至少兩個 Excel 文件進行比較");
        }

        // 設定比較索引，例如HostName 和 IP的欄位索引
        int HostnameIndex = 0;
        int ipIndex = 1;

        // 統計變數
        int sameHostnameAndIPCount = 0;
        int sameHostnameDifferentIPCount = 0;
        int differentHostnameSameIPCount = 0;
        int differentHostnameAndIPCount = 0;

        // 建立工作表
        Sheet statsSheet = workbook.createSheet("統計表");
        Sheet sheetSameHostnameAndIP = workbook.createSheet("Hostname與IP相同");
        Sheet sheetSameHostnameDifferentIP = workbook.createSheet("Hostname相同IP不同");
        Sheet sheetDifferentHostnameSameIP = workbook.createSheet("Hostname不同IP相同");
        Sheet sheetDifferentHostnameAndIP = workbook.createSheet("Hostname與IP不同");

        int rowIndexSameHostnameAndIP = 0;
        int rowIndexSameHostnameDifferentIP = 0;
        int rowIndexDifferentHostnameSameIP = 0;
        int rowIndexDifferentHostnameAndIP = 0;

        // 進行文件之間的比較
        for (int i = 0; i < files.length; i++) {
            for (int j = i + 1; j < files.length; j++) {
                Excel file1 = files[i];
                Excel file2 = files[j];

                // 取得兩個文件的數據
                HashSet<ArrayList<String>> file1Data = file1.getSourceData();
                HashSet<ArrayList<String>> file2Data = file2.getSourceData();

                // 開始比較邏輯
                for (ArrayList<String> row1 : file1Data) {
                    String hostname1 = row1.get(HostnameIndex);
                    String ip1 = row1.get(ipIndex);

                    boolean foundSameHostnameAndIP = false;
                    boolean foundSameHostnameDifferentIP = false;
                    boolean foundDifferentHostnameSameIP = false;

                    for (ArrayList<String> row2 : file2Data) {
                        String hostname2 = row2.get(HostnameIndex);
                        String ip2 = row2.get(ipIndex);

                        if (hostname1.equals(hostname2) && ip1.equals(ip2)) {
                            foundSameHostnameAndIP = true;
                            rowIndexSameHostnameAndIP = addRow(sheetSameHostnameAndIP, rowIndexSameHostnameAndIP, row1, row2);
                            sameHostnameAndIPCount++;
                        } else if (hostname1.equals(hostname2) && !ip1.equals(ip2)) {
                            foundSameHostnameDifferentIP = true;
                            rowIndexSameHostnameDifferentIP = addRow(sheetSameHostnameDifferentIP, rowIndexSameHostnameDifferentIP, row1, row2);
                            sameHostnameDifferentIPCount++;
                        } else if (!hostname1.equals(hostname2) && ip1.equals(ip2)) {
                            foundDifferentHostnameSameIP = true;
                            rowIndexDifferentHostnameSameIP = addRow(sheetDifferentHostnameSameIP, rowIndexDifferentHostnameSameIP, row1, row2);
                            differentHostnameSameIPCount++;
                        }
                    }

                    if (!foundSameHostnameAndIP && !foundSameHostnameDifferentIP && !foundDifferentHostnameSameIP) {
                        rowIndexDifferentHostnameAndIP = addRow(sheetDifferentHostnameAndIP, rowIndexDifferentHostnameAndIP, row1, null);
                        differentHostnameAndIPCount++;
                    }
                }
            }
        }

        // 添加統計結果
        addStatistics(statsSheet, sameHostnameAndIPCount, sameHostnameDifferentIPCount, differentHostnameSameIPCount, differentHostnameAndIPCount);

        // 寫入 Excel 文件
        try (FileOutputStream fileOut = new FileOutputStream("resources\\outputdata\\CompareResult.xlsx")) {
            workbook.write(fileOut);
            System.out.println("已成功輸出Excel檔案文件！");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 添加統計結果到統計表
    private void addStatistics(Sheet statsSheet, int sameHostnameAndIPCount, int sameHostnameDifferentIPCount, int differentHostnameSameIPCount, int differentHostnameAndIPCount) {
        Row header = statsSheet.createRow(0);
        header.createCell(0).setCellValue("統計項目");
        header.createCell(1).setCellValue("數量");

        Row row1 = statsSheet.createRow(1);
        row1.createCell(0).setCellValue("hostname與ip相同");
        row1.createCell(1).setCellValue(sameHostnameAndIPCount);

        Row row2 = statsSheet.createRow(2);
        row2.createCell(0).setCellValue("hostname相同ip不同");
        row2.createCell(1).setCellValue(sameHostnameDifferentIPCount);

        Row row3 = statsSheet.createRow(3);
        row3.createCell(0).setCellValue("hostname不同ip相同");
        row3.createCell(1).setCellValue(differentHostnameSameIPCount);

        Row row4 = statsSheet.createRow(4);
        row4.createCell(0).setCellValue("hostname與ip不同");
        row4.createCell(1).setCellValue(differentHostnameAndIPCount);
    }

    // 將比較結果加入至指定的工作表
    private int addRow(Sheet sheet, int rowIndex, ArrayList<String> row1, ArrayList<String> row2) {
        Row row = sheet.createRow(rowIndex++);
        int colIndex = 0;

        // 添加第一個文件的數據
        for (String cellValue : row1) {
            row.createCell(colIndex++).setCellValue(cellValue);
        }

        // 如果第二個文件有數據，添加第二個文件的數據
        if (row2 != null) {
            for (String cellValue : row2) {
                row.createCell(colIndex++).setCellValue(cellValue);
            }
        }
        return rowIndex;
    }
}
