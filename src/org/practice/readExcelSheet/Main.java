package org.practice.readExcelSheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;
import java.util.concurrent.ConcurrentHashMap;

public class Main {

    public static final String file_path = "C:\\Users\\spande67\\Documents\\SampleNameSheet.xlsx";
    public static void main(String[] args)
    {

        System.out.println("Hello World!");

        while(true){

            System.out.println("Choose the below option to perform different task on an excel sheet!");
            System.out.println("1.Create sheet\n" +
                    "2. Read sheet\n" +
                    "3.Update sheet\n" +
                    "4.Delete sheet data\n");
            Scanner in = new Scanner(System.in);
            int ch = in.nextInt();

            switch (ch){

                case 1:
                    createExcelSheet();
                    break;

                case 2:
                    readExcelSheet();
                    break;

                case 3:
                    updateExcelSheet();
                    break;

                case 4:
                    deleteExcelSheet();
                    break;

                default:
                    System.out.println("wrong choice.");

            }

            System.out.println("Want to exit: 0");
            int ex = in.nextInt();
            if(ex==0){
                System.out.println("Exiting...");
                break;
            }

        }




    }

    public static void deleteExcelSheet(){

        try{
            FileInputStream fileInputStream = new FileInputStream(file_path);
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            System.out.println("Deletes a particular row:");
            Sheet sheet = workbook.getSheetAt(0);
            sheet.removeRow(sheet.getRow(3));



//            System.out.println("Delete the entire sheet");
//            workbook.removeSheetAt(0);

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(file_path);
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.close();
        }
        catch ( FileNotFoundException e){
            e.printStackTrace();
        }
        catch (IOException e){
            e.printStackTrace();
        }



    }

    public static void updateExcelSheet(){

        try {
            FileInputStream fileInputStream = new FileInputStream(file_path);
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            Sheet sheet = workbook.getSheetAt(0);

            System.out.println("appends and updates new rows in the sheet");
            ConcurrentHashMap<Integer,String> concurrentHashMap = new ConcurrentHashMap<Integer,String>();
            concurrentHashMap.put(4,"Akash");
            concurrentHashMap.put(5,"Sagar");
            concurrentHashMap.put(6,"Rahul");


            int rowindex = sheet.getLastRowNum();

            Iterator<Integer>  itr = concurrentHashMap.keySet().iterator();
            Iterator<String>  itr2 = concurrentHashMap.values().iterator();
            //  System.out.println(itr.);
            while(itr.hasNext()){
                Row row = sheet.createRow(++rowindex);
                int cellnum =0;

//            System.out.println("ejknk");
                row.createCell(cellnum++).setCellValue(itr.next().toString());
                row.createCell(cellnum).setCellValue(itr2.next());

            }

            System.out.println("updates a row 3 with cell value: Shraddha to a new cell value as: Shivi");
            Cell  cell = sheet.getRow(2).getCell(1);
            cell.setCellValue("Shivi");
            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(file_path);
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.close();
        }
        catch (FileNotFoundException e ){
            e.printStackTrace();
        }
        catch (IOException e){
            e.printStackTrace();
        }
    }

    public static void createExcelSheet(){

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet xssfSheet = workbook.createSheet("Sample excel sheet");

        ConcurrentHashMap<Integer,String> concurrentHashMap = new ConcurrentHashMap<Integer,String>();
        concurrentHashMap.put(1,"Name");
        concurrentHashMap.put(2,"Shubhangi");
        concurrentHashMap.put(3,"Shraddha");

        int rowindex =0;

        Iterator<Integer>  itr = concurrentHashMap.keySet().iterator();
        Iterator<String>  itr2 = concurrentHashMap.values().iterator();
        //  System.out.println(itr.);
        while(itr.hasNext()){
            Row row = xssfSheet.createRow(rowindex++);
            int cellnum =0;

//            System.out.println("ejknk");
            row.createCell(cellnum++).setCellValue(itr.next().toString());
            row.createCell(cellnum).setCellValue(itr2.next());

        }

        try {
            FileOutputStream fileOutputStream  = new FileOutputStream(file_path);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            System.out.println("file successfully written to : "+ file_path);

        } catch (FileNotFoundException e){
            e.printStackTrace();
        }
        catch (IOException e){
            e.printStackTrace();

        }
    }

    public static void readExcelSheet(){

        try {
            FileInputStream fileInputStream = new FileInputStream(file_path);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while(rowIterator.hasNext()){
                Row row = rowIterator.next();
                Iterator<Cell> celIterator = row.cellIterator();

                while(celIterator.hasNext()){
                    Cell cell = celIterator.next();
                    System.out.print(cell.getStringCellValue()+" ");
                     }

                System.out.println("");
            }

            System.out.println("the contents of the excel sheet are read and printed successfully!");

        }
        catch (FileNotFoundException e){
            e.printStackTrace();
        }
        catch (IOException e){
            e.printStackTrace();
        }

    }
}
