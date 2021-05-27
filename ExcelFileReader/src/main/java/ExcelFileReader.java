import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileReader {

    public static List<ExternalDataObject> excelData;

    public ExcelFileReader() throws IOException {
        readExcelFile();
    }

    public void readExcelFile() throws IOException {
        //create file
        File file = new File("src/main/resources/exampleData.xls");

        //create input stream
        FileInputStream inputStream = new FileInputStream(file);

        // String fileExtensionName = fileName.substring(fileName.indexOf("."));
        Workbook externalData = new HSSFWorkbook(inputStream);

        //get first sheet By index
        Sheet sheet = externalData.getSheetAt(0);

        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
        excelData = new ArrayList<>();

        for (int i = 0; i < rowCount + 1; i++) {
            Row row = sheet.getRow(i);
            ExternalDataObject item = new ExternalDataObject();
            for (int j = 0; j < row.getLastCellNum(); j++) {

                switch (j) {
                    case 0:
                        item.browser = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 1:
                        item.id_number = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 2:
                        item.last_6_digits = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 3:
                        item.link = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 4:
                        item.firstName = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 5:
                        item.surname = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 6:
                        item.id_type = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 7:
                        item.passportNum = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 8:
                        item.nominee_id_number = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 9:
                        item.gender = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 10:
                        item.yearOfBirth = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 11:
                        item.monthOfBirth = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 12:
                        item.dayOfBirth = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 13:
                        item.nationality = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 14:
                        item.issuing_country = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 15:
                        item.contact_country = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 16:
                        item.countryOfResidence = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 17:
                        item.countryOfBirth = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 18:
                        item.citizenship = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 19:
                        item.visa_type = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 20:
                        item.visa_number = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 21:
                        item.vi_year = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 22:
                        item.vi_month = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 23:
                        item.vi_day = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 24:
                        item.ve_year = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 25:
                        item.ve_month = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 26:
                        item.ve_day = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 27:
                        item.contact_type = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 28:
                        item.contact_number = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 29:
                        item.preferred_comms = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 30:
                        item.house_number = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 31:
                        item.street_num = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 32:
                        item.suburb = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 33:
                        item.city = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 34:
                        item.province = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 35:
                        item.postal_code = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 36:
                        item.personOfInterest = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 37:
                        item.relatedToPIP = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 38:
                        item.relation2PIP = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 39:
                        item.pipName = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                    case 40:
                        item.pipSurname = row.getCell(j).getStringCellValue().replace("\"","");
                        break;
                }
            }
            excelData.add(item);
        }
    }

    public static Properties readPropsFile() throws IOException {
        Properties config = new Properties();
        config.load(new FileInputStream("C:\\Users\\a237902\\Documents\\test-automation\\CardAddOnAutamationTest\\config.properties"));
        return config;
    }

    public List<ExternalDataObject> getExcelData(){
        return excelData;
    }
}
