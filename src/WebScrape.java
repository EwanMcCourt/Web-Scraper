
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WebScrape {



    public static void main(String[] args) throws IOException {


        try {

            //Creates excel file
            final String filename = "April2021.xlsx";  //Change as appropriate
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("FirstSheet");



            //Creates headings
            Row rowhead = sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Winnings");
            rowhead.createCell(1).setCellValue("Account Number");
            rowhead.createCell(2).setCellValue("Holdings");
            rowhead.createCell(3).setCellValue("Area");

            //Connects to NS&I website and creates arrays for each header
            Document doc = Jsoup.connect("https://www.nsandi.com/prize-checker/winners").timeout(6000).get();
            Elements body = doc.select("tbody");
            ArrayList<String> firstArray = new ArrayList<String>();
            firstArray.add("Winnings");
            ArrayList<String> secondArray = new ArrayList<String>();
            secondArray.add("Account Number");
            ArrayList<String> thirdArray = new ArrayList<String>();
            thirdArray.add("Holdings");
            ArrayList<String> fourthArray = new ArrayList<String>();
            fourthArray.add("Area");

            //Loop through each row and add values to arrays
            for(Element e : body.select("tr")) {

                String ValueAndNo = e.select("td:nth-of-type(1)").text().trim();
                String [] arr1 = ValueAndNo.split(" ", 2);

                firstArray.add(arr1[0]);
                secondArray.add(arr1[1]);

                String HoldingAndArea = e.select("td:nth-of-type(3)").text().trim();
                String [] arr2 = HoldingAndArea.split(" ", 2);
                thirdArray.add(arr2[0]);
                fourthArray.add(arr2[1]);

            }


            //Loops through all arrays and separates their values into the appropriate cells
            XSSFRow row = sheet.createRow((short)1);
            for(int j = 1; j < firstArray.size(); j++)
            {
                row = sheet.getRow(j);
                if (row == null) {
                    row = sheet.createRow(j);
                }
                Cell cell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(firstArray.get(j));

            }


            for(int j = 1; j < secondArray.size(); j++)
            {
                row = sheet.getRow(j);
                if (row == null) {
                    row = sheet.createRow(j);
                }
                Cell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(secondArray.get(j));
            }
            for(int j = 1; j < thirdArray.size(); j++)
            {
                row = sheet.getRow(j);
                if (row == null) {
                    row = sheet.createRow(j);
                }
                Cell cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(thirdArray.get(j));
            }
            for(int j = 1; j < fourthArray.size(); j++)
            {
                row = sheet.getRow(j);
                if (row == null) {
                    row = sheet.createRow(j);
                }
                Cell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(fourthArray.get(j));
            }

            //Writes to excel file
            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Your excel file has been generated!");
        }
        catch ( Exception ex ) {
            System.out.println(ex);
        }



    }


}


