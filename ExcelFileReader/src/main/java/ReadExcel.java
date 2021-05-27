import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) throws IOException {
        ExcelFileReader read = new ExcelFileReader();
        String x = read.getExcelData().get(1).browser;
        System.out.println(x);
    }
}
