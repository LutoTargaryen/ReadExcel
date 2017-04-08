package com.lutotargaryen.poi.readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
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

public class ReadExcel {

	//返回状态
	//无效的文件路径
	public static final int INVALIDFILEPATH = -1;
	//文件类型错误
	public static final int FILETYPEERROR = -2;
	//列名不一致
	public static final int COLUMNINCONSISTENCY = -3;
	//数据库中没有相对应的表
	public static final int HAVENOTABLE = -4;
		
	
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
	//是否创建表，默认为false，如果设置为true，则数据中没有excel文件名的表时
	//创建一个以excel文件名为表名的表（字段类型为varchar）
	//private boolean isCreat;
	

	// 使用volatile关键字保证多线程访问时readExcel的可见性
	private volatile static ReadExcel readExcel;
	/**
	 * 创建ReadExcel对象
	 * 如果不设置UserName和Password则默认数据库连接用户为root,数据库连接密码也为root
	 * @param Driver 数据库驱动类路径
	 * @param DBUrl 数据库连接url
	 * @return
	 * @throws RepeatCreateObject 
	 */
	public static ReadExcel newInstance(String Driver,String DBUrl) throws RepeatCreateObject{
		if(readExcel == null){
			synchronized (ReadExcel.class) {
				if(readExcel == null){
					readExcel = new ReadExcel();
					JdbcUtil.setDRIVERCLASS(Driver);
					JdbcUtil.setURL(DBUrl);
					JdbcUtil.setUSERNAME("root");
					JdbcUtil.setPASSWROD("root");
				}
			}
			
		}else{
			throw new RepeatCreateObject("This class can have only one instance object, and you have created an instance object");
		}
		return readExcel;
	}
	/**
	 * 创建ReadExcel对象
	 * @param Driver 数据库驱动类路径
	 * @param DBUrl	数据库连接url
	 * @param UserName	数据库连接用户名
	 * @param PassWord	数据库连接密码
	 * @return
	 * @throws RepeatCreateObject 
	 */
	public static ReadExcel newInstance(String Driver,String DBUrl,String UserName,String PassWord) throws RepeatCreateObject{
		if(readExcel == null){
			synchronized (ReadExcel.class) {
				if(readExcel == null){
					readExcel = new ReadExcel();
					JdbcUtil.setDRIVERCLASS(Driver);
					JdbcUtil.setURL(DBUrl);
					JdbcUtil.setUSERNAME(UserName);
					JdbcUtil.setPASSWROD(PassWord);
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
					
				}else{
					System.out.println("列名不一致");
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
	 * 判断数据库中的字段是否和excel表中的字段一致
	 * @param excelFileds
	 * @param dBFileds
	 * @return
	 */
	private boolean getConsistent(List<String> excelFileds, List<Map<String, Object>> DBFileds) {
		for(Map<String,Object> map : DBFileds){
			System.out.println(map);
		}
		if(excelFileds.size() != DBFileds.size()){
			return false;
		}
		return false;
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
     * 获取excel表中的字段
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
