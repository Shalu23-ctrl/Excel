# Excel
#Data insertion through excel in Database
#I'm reading the Data from excel sheet using java POI and need to insert into database.I have id(numeric),name(String) and balance(double) ,actionperformed(String) from excel sheet as well as first fields are column header.

#This is the code that I wrote for the same operation. Please note that I have also created an Employee java class with all these 4 properties: id,name,balance,action performed with their getter-setters.

import ExcelData.Account;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;
import java.util.Iterator;

public class ExcelToDatabase {
    XSSFRow row;
    Account acc = new Account();

    public static void main(String[] args) throws IOException {

        String fileName = "/home/user/Documents/Account_Info.xlsx";
        ExcelToDatabase Acc2 = new ExcelToDatabase();
       Acc2.readFile(fileName);
    }

    public void readFile(String fileName) throws FileNotFoundException, IOException {
        FileInputStream fis;
        try {
            System.out.println("------READING THE SPREADSHEET-------");
            fis = new FileInputStream(fileName);
            XSSFWorkbook workbookRead = new XSSFWorkbook(fis);
            XSSFSheet spreadsheetRead = workbookRead.getSheetAt(0);


            Iterator<Row> rowIterator = spreadsheetRead.iterator();
            Row headerRow= rowIterator.next(); // to remove header part from excel sheet
            int count =0;
            while (rowIterator.hasNext()) {
                row = (XSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getColumnIndex()) {
                        case 0:
                            if(cell.getCellType().equals(CellType.NUMERIC)) {
                                System.out.print(
                                        cell.getNumericCellValue()+ " \t\t");
                            }
                            else if(cell.getCellType().equals(CellType.STRING)){
                                System.out.print(
                                        cell.getStringCellValue()+ " \t\t");
                            }
                            break;
                        case 1:
                            System.out.print(
                                    cell.getStringCellValue() + " \t\t");
                           break;
                        case 2:
                            if(cell.getCellType().equals(CellType.NUMERIC)) {
                                System.out.print(
                                        cell.getNumericCellValue()+ " \t\t");
                            } else if(cell.getCellType().equals(CellType.STRING)){
                                System.out.print(
                                        cell.getStringCellValue()+ " \t\t");
                            }
                            break;
                        case 3:
                            if(cell.getCellType().equals(CellType.NUMERIC))
                             System.out.print(
                                    cell.getNumericCellValue()+ " \t\t");
                            break;

                     case 4:
                        System.out.print(
                                  cell.getStringCellValue() + " \t\t");
break;
                    }
                    System.out.println();

                }
                if(count >0) {
//                    if (row.getCell(0).getCellType().equals(CellType.NUMERIC)) {
//                        acc.id = Integer.parseInt(row.getCell(0).getStringCellValue());
//                    } else {
//                        acc.id = Integer.parseInt(row.getCell(0).getStringCellValue());
//                    }

                    acc.name = row.getCell(1).getStringCellValue();
                    if (row.getCell(2).getCellType().equals(CellType.NUMERIC)) {
                        acc.balance = Double.parseDouble(String.valueOf(row.getCell(2).getNumericCellValue()));
                    } else {
                        acc.balance = Double.parseDouble(row.getCell(2).getStringCellValue());
                    }
                    acc.Actionperformed = row.getCell(3).getStringCellValue();

                    InsertRowInDB(acc.id, acc.name, acc.balance, acc.Actionperformed);
                }
                count ++;
            }
                System.out.println();
            }

        catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Values Inserted Successfully");
    }

    public void InsertRowInDB(int id, String name, Double balance, String ActionPerformed) {

        try {

            Class.forName("org.postgresql.Driver");
            Connection c;
            c = DriverManager
                    .getConnection("jdbc:postgresql://localhost:5432/Account_info",
                            "postgres", "whizzard@123");
            System.out.println("Opened database successfully");

            PreparedStatement ps = null;

            Statement stmt = c.createStatement();

            String sql = "INSERT INTO INFO (NAME,BALANCE,ACTIONPERFORMED) VALUES(?,?,?)";
            ps = c.prepareStatement(sql);
           // ps.setInt(1, id);
            ps.setString(1, name);
            ps.setDouble(2, balance);
            ps.setString(3, ActionPerformed);
            ps.executeUpdate();
            c.close();

        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
    }

}




