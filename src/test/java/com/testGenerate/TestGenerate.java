package com.testGenerate;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Locale;
import java.util.Scanner;


import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class TestGenerate {
    public static void main(String[] args) throws  java.io.IOException {
        XSSFCell cell;
        FileInputStream fis = new FileInputStream("C:\\Users\\ADMIN\\IdeaProjects\\TestDataGenerate\\TestData.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        Locale local = new Locale("en-IND");
        Faker fake = new Faker(local);
        System.out.println("Enter the number of rows of data you want to print !!");
        Scanner sc = new Scanner(System.in);
        int rowcount = sc.nextInt();
        System.out.println("Enter the number of columns of data you want to print !!");
        int colcount = sc.nextInt();

        if (colcount >= 1 && colcount <= 6) {
            int counter = 1;
            for (int i = 0; i <= rowcount-1; i++) {
                XSSFRow row = sheet.createRow(i);

                for (int j = 0; j <= colcount; j++) {
                    cell = row.createCell(j);
                    String columnname = "TestData" + counter;
                    cell.setCellValue(columnname);
                    if (j == 1) {
                        cell = row.createCell(j);
                        String firstname = fake.name().firstName();
                        cell.setCellValue(firstname);
                    }
                    if (j == 2) {
                        cell = row.createCell(j);
                        String lastname = fake.name().lastName();
                        cell.setCellValue(lastname);
                    }
                    if (j == 3) {
                        cell = row.createCell(j);
                        String country = fake.address().country();
                        cell.setCellValue(country);
                    }
                    if (j == 4) {
                        cell = row.createCell(j);
                        String email = fake.internet().emailAddress();
                        cell.setCellValue(email + "1");
                    }
                    if (j == 5) {
                        cell = row.createCell(j);
                        String phno = fake.phoneNumber().cellPhone();
                        cell.setCellValue(phno);
                    }
                    if (j == 6) {
                        cell = row.createCell(j);
                        String id = fake.idNumber().invalid();
                        cell.setCellValue(id);
                    }

                }
                counter+=1;

            }
            FileOutputStream fout = new FileOutputStream("C:\\Users\\ADMIN\\IdeaProjects\\TestDataGenerate\\TestData.xlsx");
            wb.write(fout);
            wb.close();
            System.out.println("Test Data Generated Successfully");
        }
        else {
            System.out.println("ColumnCount Exceeded Test Data not Generated Maximum "+colcount+" column you can print .");
            
        }
    }
}
