package test_utility;

import utilities.excel_utilities.SheetData;

public class Main {
    public static void main(String[] args) {
        String fullName = SheetData.getFullName("John");
        System.out.println(fullName);

        String name = SheetData.getOnlyName("Adam");
        System.out.println(name);

        System.out.println(SheetData.getAllInfo("Adam"));
    }
}
