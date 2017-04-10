package com.lutotargaryen.poi.readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lutotargaryen.poi.exception.RepeatCreateObject;

public class ReadExcel implements Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	/**
	 * 返回状态 无效的文件路径
	 */
	public static final int INVALIDFILEPATH = -1;
	/**
	 * 返回状态 文件类型错误
	 */
	public static final int FILETYPEERROR = -2;
	/**
	 * 返回状态 数据库中没有相对应的表
	 */
	public static final int HAVENOTABLE = -4;
	/**
	 * 返回状态 数据库中列名和excel表中列名不一致
	 */
	public static final int COLUMNINCONSISTENCY = -3;
	/**
	 * 返回状态 excel中数据类型和数据库中的数据类型不一致
	 */
	public static final int DATATYPEERROR = -5;
	/**
	 * 返回状态 导入成功
	 */
	public static final int SUCCESS = 1;
	
	//文件类型
	private String fileType;
	//文件名
	private String fileName;
	//创建需要用到的输入流
	private InputStream is;
	//创建工作簿对象
	private Workbook workBook;
	//工作簿中的sheet表数量
	private Integer numOfSheets;
	//数据库中是否存在excel文件名的表，默认不存在
	private boolean isExist;
	//列名与其数据类型的集合
	private Map<String,String> filedType = new HashMap<>();
	

	// 使用volatile关键字保证多线程访问时readExcel的可见性
	private volatile static ReadExcel readExcel;
	
	/**
	 * 创建ReadExcel对象
	 * @param connection	数据库连接对象
	 * @return
	 * @throws RepeatCreateObject
	 */
	public static ReadExcel newInstance(Connection connection) throws RepeatCreateObject{
		if(readExcel == null){
			synchronized (ReadExcel.class) {
				if(readExcel == null){
					readExcel = new ReadExcel();
					JdbcUtil.setConnnection(connection);
				}
			}
			
		}else{
			throw new RepeatCreateObject("This class can have only one instance object, and you have created an instance object");
		}
		return readExcel;
	}
	
	/**
	 * 导入excel中的数据到mysql数据库中
	 * @param path	excel表路径
	 * @return
	 * @throws IOException 
	 */
	public int readExcelToMysql(String path) throws IOException{
		//判断路径是否为空
		if(path == null || path.trim().equals("")){
			return INVALIDFILEPATH;
		}
		this.is = new FileInputStream(path);
		//得到文件类型
		this.fileType = getFileType(path);
		//创建工作表对象
		this.workBook = getWorkBook(this.fileType,this.is);
		if(workBook == null){
			return FILETYPEERROR;
		}
		//获取文件名
		this.fileName = getFileName(path);
		//获取数据库中的表名
		List<String> tables = JdbcUtil.getTable();
		//判断数据库中是否有存在和excel文件名相同的表
		for(String table : tables){
			if(table.equalsIgnoreCase(this.fileName)){
				this.isExist = true;
				break;
			}
		}
		if(this.isExist){
			//数据库中存在表
			//得到工作簿中的sheet表数量
			this.numOfSheets = workBook.getNumberOfSheets();
			//循环工作表sheet
			for(int sheetNum = 0;sheetNum < this.numOfSheets;sheetNum++){
				//取出sheet对象
				Sheet sheet = workBook.getSheetAt(sheetNum);
				//如果sheet中的有效行数0,则该sheet中没有数据
				if(sheet.getLastRowNum() == 0){
					continue;
				}
				//获取excel表中的列
				List<String> excelFileds = getExcelFileds(sheet);
				//获取数据库中的所有列
				List<Map<String,Object>> DBFileds = JdbcUtil.getColumns(this.fileName);
				//判断数据库列明和excel表列明是否一直
				boolean isConsistent = getConsistent(excelFileds,DBFileds);
				
				if(isConsistent){
					//拼接sql插入语句
					StringBuffer sql = new StringBuffer("INSERT INTO ");
					sql.append(fileName + "(");
					sql.append(getArrays(excelFileds));
					sql.append(") VALUES (");
					for(int i = 0;i<excelFileds.size();i++){
						sql.append("?");
						if(i+1 < excelFileds.size()){
							sql.append(",");
						}
					}
					sql.append(")");
					//循环sheet中的每一行
					List<Object[]> list = new ArrayList<>();
					for(int i = 1;i<=sheet.getLastRowNum();i++){
						//获取每一行的数据
						Object[] r = null;
						
						try {
							r = getRow(sheet,i,sheet.getRow(0).getLastCellNum());
						} catch (Exception e) {
							//数据类型不对
							return DATATYPEERROR;
						}
						list.add(r);
					}
					
					int res = JdbcUtil.executeUpdate(sql.toString(), list);
					System.out.println(res == sheet.getLastRowNum());
					if(res == sheet.getLastRowNum()){
						//导入成功
						return SUCCESS;
					}
				}else{
					return COLUMNINCONSISTENCY;
				}
			}
			
		}else{
			//数据库中没有表
			return HAVENOTABLE;
		}
		return 0;
	}
	/**
	 * 获取每一行的数据
	 * @param sheet	
	 * @param i 第几行
	 * @param num 列的数量
	 * @return
	 * @throws Exception 
	 */
	private Object[] getRow(Sheet sheet,int i, int num) throws Exception{
		List<Object> rows = new ArrayList<>();
		Row r = sheet.getRow(i);
		for(int j = 0;j<num;j++){
			//判断是否为空列，如果是，跳出本次循环
			Cell cell = sheet.getRow(0).getCell(j);
			if(cell == null ){
				continue;
			}
			//获取该列在数据库中的类型
			String fType = filedType.get(cell.toString());
			Object field = null;
			//根据类型转换为相应的类型
			if(fType.equalsIgnoreCase("int")){
				Double d = Double.valueOf(r.getCell(j).toString());
				if(d * 10 > d.intValue() * 10){
					throw new Exception();
				}
				field = d.intValue();
			}else if(fType.equalsIgnoreCase("double") ){
				Double d = Double.valueOf(r.getCell(j).toString());
				field =d;
			}else if(fType.equalsIgnoreCase("float")){
				Double d = Double.valueOf(r.getCell(j).toString());
				field =d.floatValue();
			}else if(fType.equalsIgnoreCase("date")){
				field = DateUtil.StringToDate(r.getCell(j).toString());
			}else{
				field = String.valueOf(r.getCell(j));
			}
			rows.add(field);
		}
		return rows.toArray(new Object[rows.size()]);
	}
	/**
	 * 
	 * @param list
	 * @return
	 */
	private String getArrays(List<String> list) {
		StringBuffer sb = new StringBuffer();
		for(int i = 0;i<list.size();i++){
			sb.append(list.get(i));
			if(i+1 < list.size()){
				sb.append(",");
			}
		}
		return sb.toString();
	}
	/**
	 * 判断数据库中的字段是否和excel表中的字段一致
	 * @param excelFileds
	 * @param dBFileds
	 * @return
	 */
	private boolean getConsistent(List<String> excelFileds, List<Map<String, Object>> DBFileds) {
		if(excelFileds.size() != DBFileds.size()){
			return false;
		}
		boolean t = false;
		for (Map<String, Object> map : DBFileds) {
			String DBFiled = (String) map.get("COLUMN_NAME");
			for(String excelFiled : excelFileds){
				if(excelFiled.trim().equalsIgnoreCase(DBFiled)){
					//把列名还有该列名的数据类型存储到map集合中
					filedType.put(excelFiled, (String)map.get("TYPE_NAME"));
					t = true;
					break;
				}else{
					t = false;
				}
			}
		}
		if(!t){
			return false;
		}
		return true;
	}
	/**
	 * 获取文件名
	 * @param path
	 * @return
	 */
	private String getFileName(String path) {
		String fileName = null;
		File file = new File(path);
		fileName = (String) file.getName().subSequence(0,file.getName().indexOf("."));
		return fileName;
	}
	
	/**
	 * 根据路劲获取文件类型
	 * @param path
	 * @return
	 */
	private String getFileType(String path) {
		String fileType = path.substring(path.indexOf(".") + 1);
		return fileType;
	}
	/**
	 * 根据文件是xls还是xlsx创建不同的对象
	 * @param fileType
	 * @return
	 * @throws IOException 
	 */
	private Workbook getWorkBook(String fileType,InputStream is) throws IOException {
		Workbook workBook = null;
		//判断文件是什么类型
		if("xlsx".equalsIgnoreCase(fileType)){
			//xlsx文件用XSSFWorkbook创建对象
			workBook = new XSSFWorkbook(is);
		}else if("xls".equalsIgnoreCase(fileType)){
			//xls文件用HSSFWorkbook创建对象
			workBook = new HSSFWorkbook(is);
		}
		return workBook;
	}
	/**
     * 获取excel表中的列名
     * @param hssfSheet HSSFSheet对象(excel中的sheet)
     * @return
     */
    private List<String> getExcelFileds(Sheet sheet) {
    	List<String> excelFileds = null;
    	//获取sheet表中的第一行
    	Row row = sheet.getRow(0);
		if(row != null){
			excelFileds = new ArrayList<>();
			//遍历一行中的所有列
			Iterator<Cell> excelRow = row.cellIterator();
			while(excelRow.hasNext()){
				String excelFiled = excelRow.next().getStringCellValue();
				excelFileds.add(excelFiled);
			}
		}
		return excelFileds;
	}
}

