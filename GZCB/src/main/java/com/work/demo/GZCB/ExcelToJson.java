package com.work.demo.GZCB;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
public class ExcelToJson 
{
    public static void main( String[] args )
    {
    	 XSSFWorkbook book;
         XSSFSheet sheet;
         JSONArray jsons;
         XSSFRow row;

         try {
             InputStream is = new FileInputStream(new File( "C:\\Users\\dell\\Desktop\\file\\data.xlsx"));

             book = new XSSFWorkbook(is);

             sheet = book.getSheetAt(1);

             jsons = new JSONArray();

             for(int i = 1; i < 10; i++) {
                 row = sheet.getRow(i);
                 if(row != null) {
                     JSONObject json = new JSONObject();
                     //对于纯数字内容要做这一操作
                     row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                     row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                    
                     json.put("id", row.getCell(0).getStringCellValue());
                     json.put("name", row.getCell(1).getStringCellValue());
                     jsons.add(json);
                 }
             }

             System.out.println(jsons.toJSONString());
             book.close();} catch (FileNotFoundException e) {
                 // TODO 自动生成的 catch 块
                 e.printStackTrace();
             } catch (IOException e) {
                 // TODO 自动生成的 catch 块
                 e.printStackTrace();
             }

         }
}
