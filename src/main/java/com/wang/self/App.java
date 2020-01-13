package com.wang.self;





import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Properties;
import java.util.stream.Stream;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws IOException {


        Properties prop = new Properties();
        try {
            prop.load(new BufferedReader(new FileReader("D:/exceltools.properties")));

        } catch (IOException e) {
            e.printStackTrace();
        }

        String suffix= prop.getProperty("filesuffix");
        Workbook newWB = null;
        if("xlsx".equals(suffix)){
            newWB = new XSSFWorkbook();
        }else{
            newWB = new HSSFWorkbook();
        }



       String type=  prop.getProperty("type");
       if("unionfile".equals(type)){
         String filenames = prop.getProperty("filedirectory");
           System.out.println(filenames);

           String fileName = filenames+"//workbook."+suffix;
           Path p = Paths.get(fileName);
           System.out.println(p);
           if(Files.exists(p)){
               System.out.println("存在");
               Files.delete(p);
           }

       //  String[] filenameArray = filenames.split(",");
        Sheet newSheet =  newWB.createSheet();

           File folder = new File(filenames);
           File[] fileArray = folder.listFiles();
         int rowline =0;
           DataFormatter formatter = new DataFormatter();
           boolean addHeader=false;
         for(File filename:fileArray){
             try (InputStream inp = new FileInputStream(filename)) {
                 Workbook wbtemp= null;
                 if("xlsx".equals(suffix)){
                     wbtemp = new XSSFWorkbook(inp);
                 }else{
                     wbtemp = new HSSFWorkbook(inp);
                 }


                 Sheet sheettemp = wbtemp.getSheetAt(0);

                 for(Row row:sheettemp){
                     boolean hasvalue = false;
                     if(addHeader){
                         if(row.getRowNum()==0){
                             continue;
                         }
                     }
                     Row newRow = newSheet.createRow(rowline);
                     for (int cellCount=0;cellCount<row.getLastCellNum();cellCount++) {
                         Cell oldCell = row.getCell(cellCount);
                         if(oldCell==null){
                             continue;
                         }
                         Cell newCell = newRow.createCell(cellCount);
                         String text = formatter.formatCellValue(oldCell);
                         if(!text.isEmpty()){
                             hasvalue=true;
                         }
                         switch (oldCell.getCellType()) {
                             case STRING:
                                 newCell.setCellValue (oldCell.getRichStringCellValue().getString());
                                 break;
                             case NUMERIC:
                                 if (DateUtil.isCellDateFormatted(oldCell)) {
                                    Date a = oldCell.getDateCellValue();
                                     SimpleDateFormat spf = new SimpleDateFormat("yyyy/MM/dd");
                                      String dateString =spf.format(a);
                                     newCell.setCellValue (dateString);
                                 } else {
                                     newCell.setCellValue (oldCell.getNumericCellValue());
                                 }
                                 break;
                             case BOOLEAN:
                                 newCell.setCellValue (oldCell.getBooleanCellValue());
                                 break;
                             case FORMULA:
                                 newCell.setCellValue (oldCell.getCellFormula());
                                 break;
                             case BLANK:
                                 newCell.setCellValue("");
                                 break;
                             default:
                                 newCell.setCellValue("");
                         }
                     }
                     if(hasvalue){
                         rowline++;
                     }else{
                         newSheet.removeRow(newRow);
                     }
                 }

             } catch (FileNotFoundException e) {
                 e.printStackTrace();
             } catch (IOException e) {
                 e.printStackTrace();
             }
             addHeader = true;
         }


           try (OutputStream fileOut = new FileOutputStream(fileName)) {
               newWB.write(fileOut);
           }catch (Exception e){

           }
       }

    }



}
