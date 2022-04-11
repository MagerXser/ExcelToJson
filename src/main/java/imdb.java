import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class imdb {

    private String title;
    private String score;
    private String year;
    private String duration;
    private String rating;
    private String budget;
    private String genres;
    private String gross;
    private String director;
    private String actor1;
    private String actor2;
    private String actor3;
    private String language;
    private String country;


    public imdb() {
    }

//    public String toString() {
//        return String.format("%s - %s - %f", title, author, price);
//    }

    public static void main(String[] args) throws IOException {
        imdb imdb = new imdb();

        imdb.readSheetWithFormula("C://Users//MagerXser//Downloads//imdb_data.csv");
    }

    public static void readSheetWithFormula(String filePath)
    {
        try
        {
            FileInputStream file = new FileInputStream(new File(filePath));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type after eveluating formulae
                    //If it is formula cell, it will be evaluated otherwise no change will happen
                    switch (evaluator.evaluateInCell(cell).getCellType())
                    {
                        case
                                NUMERIC:
                            System.out.print(cell.getStringCellValue() + "tt");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "tt");
                            break;
                        case BOOLEAN:
                            //Not again
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}
