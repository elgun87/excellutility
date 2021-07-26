package utilities.excel_utilities;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utilities.ConfigurationReader;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class SheetData {
    /*
    Methods in this class will get specific data from excel sheet
     */
    static FileInputStream fileInputStream;
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;

    static {
        try {
            fileInputStream = new FileInputStream("SampleData.xlsx");
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheet(ConfigurationReader.getProperty("sheetName"));
        } catch (Exception exceptione) {
            exceptione.printStackTrace();
        }
    }

    /*
    below method will return full name of given name,if not will return null
     */
    public static String getFullName(String name) {
        int lastUsedRow = sheet.getLastRowNum();
        for (int i = 1; i <= lastUsedRow; i++) {
            String data = "";
            for (int j = 0; j < lastUsedRow - 2; j++) {
                data += " " + sheet.getRow(i).getCell(j).toString();
            }
            if (data.contains(name)) {
                data = data.trim();
                return data;
            }
        }
        return null;
    }

    /*
    below method will return only given name from excel,if not found will return null
     */
    public static String getOnlyName(String name) {
        int lastUsedRow = sheet.getLastRowNum();
        for (int i = 1; i <= lastUsedRow - 1; i++) {
            String data = "";
            for (int j = 0; j < lastUsedRow - 3; j++) {
                data += " " + sheet.getRow(i).getCell(j).toString();
            }
            if (data.contains(name)) {
                data = data.trim();
                return data;
            }
        }
        return null;
    }

    /*
    below method will accept string name and return all information of person as a string type list
    if name is not in the excel it will return empty list
     */
    public static List<String> getAllInfo(String name) {
        int lastUsedRow = sheet.getLastRowNum();
        List<String> info = new ArrayList<>();
        for (int i = 1; i <= lastUsedRow; i++) {
            if (sheet.getRow(i).getCell(0).toString().equals(name)) {
                for (int j = 0; j < lastUsedRow - 1; j++) {
                    info.add(sheet.getRow(i).getCell(j).toString());
                }
            }
        }
        return info;
    }

}