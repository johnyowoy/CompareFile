import Config.Excel;
import Config.Compare;

public class Test{
    public static void main(String[] args) {

        var file1 = new Excel();
        file1.setExcelFileName("ComparA");
        file1.setRows(new int[]{0, 1, 2});
        file1.CollectionStore();

        var file2 = new Excel();
        file2.setExcelFileName("ComparB");
        file2.setRows(new int[]{0, 1, 2});
        file2.CollectionStore();

        var compare1 = new Compare();
        compare1.setFile(new Excel[]{file1, file2});
        compare1.compareFiles();

    }
}