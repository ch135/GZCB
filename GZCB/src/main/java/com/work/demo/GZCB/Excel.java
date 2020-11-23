package com.work.demo.GZCB;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

/**
* @author scholarly
* @version 2020年11月22日 上午10:14:39
* 
* <h5>Excel 工具类</h5>
*/
public class Excel {
	
	private String path;
	private String savePath;
	private int index = 0;
	private String name;
	
	public Excel() {}
	
	/**
	 * 
	 * @Title: Excel
	 * @Desc: 构造函数 
	 * @param path 路径
	 * @param index 表格编号
	 * @param name 更改目标
	 * @param savePath 文件保存路径
	 */
	public Excel(String path, int index, String name,String savePath) {
		this.path = path;
		this.index = index;
		this.name = name;
		this.savePath = savePath;
	}
	
	/**
	 * 
	 * @Title: changeValue 
	 * @Desc: 更改Excel数据
	 */
	public void changeValue() {
		InputStream inputfile;
		OutputStream outputfile;
		XSSFWorkbook work ;
		XSSFSheet sheet;
		int number = 0;
		Row frontrow ;
		
		try {
			inputfile = getInputStream(this.path);
			
			work = new XSSFWorkbook(inputfile);
			
			sheet = work.getSheetAt(this.index);
			
			for(Row row:sheet) {
				if(this.name.equals(getCellValue(row.getCell(1)))) {
					frontrow = sheet.getRow(number-1);
					
					String value = getCellValue(frontrow.getCell(1));
					
					setCellValue(row.getCell(1), value);
				}
				
				number++;
			}
			
			outputfile = getOutputStream(this.path, false);
			
			work.write(outputfile);
			
			outputfile.close();
			
			inputfile.close();
			
		} catch (Exception e) {
			// TODO: handle exception
		}
	}
	
	/**
	 * 
	 * @Title: toJson 
	 * @Desc: excel转化为JSON
	 */
	public void toJson() {
		XSSFWorkbook book;
        XSSFSheet sheet;
        JSONArray provinceArray = new JSONArray();
        JSONArray cityArray = new JSONArray();
        JSONArray regionsArray =  new JSONArray();
        XSSFRow row;
        InputStream inputfile;

        try {
            inputfile = getInputStream(this.path);

            book = new XSSFWorkbook(inputfile);

            sheet = book.getSheetAt(this.index);
            
            for(int i=sheet.getLastRowNum();i>=0;i--) {
	           	 row = sheet.getRow(i);
	           	 if(row!=null) {
	           		 String id = getCellValue(row.getCell(0));
	           		 
	           		 if("00".equals(id.substring(2, 4))) {
	       				 provinceArray.add(getJSONObject(row, cityArray));
	       				 cityArray = new JSONArray();
	           		 }else if("00".equals(id.substring(4, 6))) {
	           			 cityArray.add(getJSONObject(row));
	           		 }else {
	           			 regionsArray.add(getJSONObject(row, "00"));
	           		 }
	           	 }
            }
            writeToJs("var province ="+provinceArray.toString(), true);
            writeToJs("var area ="+regionsArray.toString(), true);
            
            book.close();
            inputfile.close();
        } catch (FileNotFoundException e) {
            // TODO 自动生成的 catch 块
            e.printStackTrace();
        } catch (IOException e) {
            // TODO 自动生成的 catch 块
            e.printStackTrace();
        }
	}
	
	/**
	 * 
	 * @Title: CellRemove 
	 * @Desc: 删除Excel某行
	 * @param number 单元格坐标
	 */
	public void RowRemove(int number) {
		InputStream is;
		OutputStream out;
		XSSFWorkbook work;
		XSSFSheet sheet;
		XSSFRow row;
		 try {
			 is = getInputStream(this.path);
			 work = new XSSFWorkbook(is);
			 sheet = work.getSheetAt(this.index);
			 row = sheet.getRow(number);
			 sheet.removeRow(row);
			 //sheet.shiftRows(startRow, endRow, n); 移除几行
			 out = getOutputStream(this.path, false);
			 work.write(out);
			 
			 work.close();
			 out.close();
			 is.close();
        } catch (Exception e) { 
            e.printStackTrace();
        }
	}
	public JSONObject getJSONObject(XSSFRow row) {
    	JSONObject json = new JSONObject();
        
        json.put("id", getCellValue(row.getCell(0)));
        json.put("name", getCellValue(row.getCell(1)));
        return json;
    }
    
    public JSONObject getJSONObject(XSSFRow row,JSONArray cityArray) {
    	JSONObject json = getJSONObject(row);
    	
    	json.put("city", cityArray);
    	return json;
    }
    
   public JSONObject getJSONObject(XSSFRow row,String pid) {
	   JSONObject json = getJSONObject(row);
	   
	   json.put("pid", json.get("id").toString().substring(0,4)+pid);
	   return json;
   }
   
   public void writeToJs(String str,boolean append){
	    OutputStream out = null;
	    try{
	    	out = getOutputStream(this.savePath, append);
	    	byte[] buff=str.getBytes();
	        out.write(buff);
	        out.flush();
	        out.close();
	    }catch(Exception e){
	        e.printStackTrace();
	    }
   }
   
	/**
	 * 
	 * @Title: getInputStream 
	 * @Desc: 获取输入信息流 
	 * @param path
	 * @return
	 * @throws IOException 
	 */
	public InputStream getInputStream(String path) throws IOException {
		
		return new FileInputStream(path);
	}
	
	/**
	 * 
	 * @Title: getOutputStream 
	 * @Desc: TODO 
	 * @param path	文件路径，没有时自动创建
	 * @param append 数据合并还是覆盖 
	 * @return
	 * @throws IOException
	 */
	public OutputStream getOutputStream(String path,boolean append) throws IOException {

		return new FileOutputStream(path,append);
	}
	
	/**
	 * 
	 * @Title: getCellValue 
	 * @Desc: 获取单元格的值
	 * @param cell
	 * @return
	 */
	public String getCellValue(Cell cell) {
		cell.setCellType(Cell.CELL_TYPE_STRING);
		
		return cell.getStringCellValue();
	}
	
	/**
	 * 
	 * @Title: setCellValue 
	 * @Desc: 设置单元格的值
	 * @param cell
	 * @param value
	 */
	public void setCellValue(Cell cell,String value) {
		cell.setCellType(Cell.CELL_TYPE_STRING);
		
		cell.setCellValue(value);
	}
	
	public static void main(String[] args) {
		Excel tool = new Excel("./change.xlsx",0,"市辖区","./change.js");
		tool.RowRemove(0);
	}
}
